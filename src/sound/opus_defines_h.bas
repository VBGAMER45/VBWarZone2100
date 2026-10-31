Attribute VB_Name = "opus_defines_h"
' libopus 1.5.2 (ref/opus/include): opus_defines.h's constants (error codes, the ctl request numbers, the
' values the ctls take). The ctl macros (OPUS_SET_GAIN(x) ...) are the request constants with the value passed
' to the *_ctl functions' ByRef value (VB: a GET writes it, a SET reads it). Shared by the opus_* and celt_*
' modules (celt.h's own requests are in O2's celt modules).
Option Explicit

'/** @defgroup opus_errorcodes Error codes
' * @{
' */
'/** No error @hideinitializer*/
Public Const OPUS_OK As Long = 0
'/** One or more invalid/out of range arguments @hideinitializer*/
Public Const OPUS_BAD_ARG As Long = -1
'/** Not enough bytes allocated in the buffer @hideinitializer*/
Public Const OPUS_BUFFER_TOO_SMALL As Long = -2
'/** An internal error was detected @hideinitializer*/
Public Const OPUS_INTERNAL_ERROR As Long = -3
'/** The compressed data passed is corrupted @hideinitializer*/
Public Const OPUS_INVALID_PACKET As Long = -4
'/** Invalid/unsupported request number @hideinitializer*/
Public Const OPUS_UNIMPLEMENTED As Long = -5
'/** An encoder or decoder structure is invalid or already freed @hideinitializer*/
Public Const OPUS_INVALID_STATE As Long = -6
'/** Memory allocation has failed @hideinitializer*/
Public Const OPUS_ALLOC_FAIL As Long = -7
'/**@}*/

'/** These are the actual Encoder CTL ID numbers.
'  * They should not be used directly by applications.
'  * In general, SETs should be even and GETs should be odd.*/
Public Const OPUS_SET_APPLICATION_REQUEST As Long = 4000
Public Const OPUS_GET_APPLICATION_REQUEST As Long = 4001
Public Const OPUS_SET_BITRATE_REQUEST As Long = 4002
Public Const OPUS_GET_BITRATE_REQUEST As Long = 4003
Public Const OPUS_SET_MAX_BANDWIDTH_REQUEST As Long = 4004
Public Const OPUS_GET_MAX_BANDWIDTH_REQUEST As Long = 4005
Public Const OPUS_SET_VBR_REQUEST As Long = 4006
Public Const OPUS_GET_VBR_REQUEST As Long = 4007
Public Const OPUS_SET_BANDWIDTH_REQUEST As Long = 4008
Public Const OPUS_GET_BANDWIDTH_REQUEST As Long = 4009
Public Const OPUS_SET_COMPLEXITY_REQUEST As Long = 4010
Public Const OPUS_GET_COMPLEXITY_REQUEST As Long = 4011
Public Const OPUS_SET_INBAND_FEC_REQUEST As Long = 4012
Public Const OPUS_GET_INBAND_FEC_REQUEST As Long = 4013
Public Const OPUS_SET_PACKET_LOSS_PERC_REQUEST As Long = 4014
Public Const OPUS_GET_PACKET_LOSS_PERC_REQUEST As Long = 4015
Public Const OPUS_SET_DTX_REQUEST As Long = 4016
Public Const OPUS_GET_DTX_REQUEST As Long = 4017
Public Const OPUS_SET_VBR_CONSTRAINT_REQUEST As Long = 4020
Public Const OPUS_GET_VBR_CONSTRAINT_REQUEST As Long = 4021
Public Const OPUS_SET_FORCE_CHANNELS_REQUEST As Long = 4022
Public Const OPUS_GET_FORCE_CHANNELS_REQUEST As Long = 4023
Public Const OPUS_SET_SIGNAL_REQUEST As Long = 4024
Public Const OPUS_GET_SIGNAL_REQUEST As Long = 4025
Public Const OPUS_GET_LOOKAHEAD_REQUEST As Long = 4027
'/* #define OPUS_RESET_STATE 4028 */
Public Const OPUS_GET_SAMPLE_RATE_REQUEST As Long = 4029
Public Const OPUS_GET_FINAL_RANGE_REQUEST As Long = 4031
Public Const OPUS_GET_PITCH_REQUEST As Long = 4033
Public Const OPUS_SET_GAIN_REQUEST As Long = 4034
Public Const OPUS_GET_GAIN_REQUEST As Long = 4045       '/* Should have been 4035 */
Public Const OPUS_SET_LSB_DEPTH_REQUEST As Long = 4036
Public Const OPUS_GET_LSB_DEPTH_REQUEST As Long = 4037
Public Const OPUS_GET_LAST_PACKET_DURATION_REQUEST As Long = 4039
Public Const OPUS_SET_EXPERT_FRAME_DURATION_REQUEST As Long = 4040
Public Const OPUS_GET_EXPERT_FRAME_DURATION_REQUEST As Long = 4041
Public Const OPUS_SET_PREDICTION_DISABLED_REQUEST As Long = 4042
Public Const OPUS_GET_PREDICTION_DISABLED_REQUEST As Long = 4043
'/* Don't use 4045, it's already taken by OPUS_GET_GAIN_REQUEST */
Public Const OPUS_SET_PHASE_INVERSION_DISABLED_REQUEST As Long = 4046
Public Const OPUS_GET_PHASE_INVERSION_DISABLED_REQUEST As Long = 4047
Public Const OPUS_GET_IN_DTX_REQUEST As Long = 4049
Public Const OPUS_SET_DRED_DURATION_REQUEST As Long = 4050
Public Const OPUS_GET_DRED_DURATION_REQUEST As Long = 4051
Public Const OPUS_SET_DNN_BLOB_REQUEST As Long = 4052

'/** @defgroup opus_ctlvalues Pre-defined values for CTL interface
'  * @see opus_genericctls, opus_encoderctls
'  * @{
'  */
'/* Values for the various encoder CTLs */
Public Const OPUS_AUTO As Long = -1000                  '/**<Auto/default setting @hideinitializer*/
Public Const OPUS_BITRATE_MAX As Long = -1              '/**<Maximum bitrate @hideinitializer*/

'/** Best for most VoIP/videoconference applications where listening quality and intelligibility matter most
' * @hideinitializer */
Public Const OPUS_APPLICATION_VOIP As Long = 2048
'/** Best for broadcast/high-fidelity application where the decoded audio should be as close as possible to the input
' * @hideinitializer */
Public Const OPUS_APPLICATION_AUDIO As Long = 2049
'/** Only use when lowest-achievable latency is what matters most. Voice-optimized modes cannot be used.
' * @hideinitializer */
Public Const OPUS_APPLICATION_RESTRICTED_LOWDELAY As Long = 2051

Public Const OPUS_SIGNAL_VOICE As Long = 3001           '/**< Signal being encoded is voice */
Public Const OPUS_SIGNAL_MUSIC As Long = 3002           '/**< Signal being encoded is music */
Public Const OPUS_BANDWIDTH_NARROWBAND As Long = 1101   '/**< 4 kHz bandpass @hideinitializer*/
Public Const OPUS_BANDWIDTH_MEDIUMBAND As Long = 1102   '/**< 6 kHz bandpass @hideinitializer*/
Public Const OPUS_BANDWIDTH_WIDEBAND As Long = 1103     '/**< 8 kHz bandpass @hideinitializer*/
Public Const OPUS_BANDWIDTH_SUPERWIDEBAND As Long = 1104    '/**<12 kHz bandpass @hideinitializer*/
Public Const OPUS_BANDWIDTH_FULLBAND As Long = 1105     '/**<20 kHz bandpass @hideinitializer*/

Public Const OPUS_FRAMESIZE_ARG As Long = 5000          '/**< Select frame size from the argument (default) */
Public Const OPUS_FRAMESIZE_2_5_MS As Long = 5001       '/**< Use 2.5 ms frames */
Public Const OPUS_FRAMESIZE_5_MS As Long = 5002         '/**< Use 5 ms frames */
Public Const OPUS_FRAMESIZE_10_MS As Long = 5003        '/**< Use 10 ms frames */
Public Const OPUS_FRAMESIZE_20_MS As Long = 5004        '/**< Use 20 ms frames */
Public Const OPUS_FRAMESIZE_40_MS As Long = 5005        '/**< Use 40 ms frames */
Public Const OPUS_FRAMESIZE_60_MS As Long = 5006        '/**< Use 60 ms frames */
Public Const OPUS_FRAMESIZE_80_MS As Long = 5007        '/**< Use 80 ms frames */
Public Const OPUS_FRAMESIZE_100_MS As Long = 5008       '/**< Use 100 ms frames */
Public Const OPUS_FRAMESIZE_120_MS As Long = 5009       '/**< Use 120 ms frames */
'/**@}*/

'/** Resets the codec state to be equivalent to a freshly initialized state.
'  * This should be called when switching streams in order to prevent
'  * the back to back decoding from giving different results from
'  * one at a time decoding.
'  * @hideinitializer */
Public Const OPUS_RESET_STATE As Long = 4028
