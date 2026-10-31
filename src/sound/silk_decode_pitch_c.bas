Attribute VB_Name = "silk_decode_pitch_c"
' libopus 1.5.2 (ref/opus/silk): decode_pitch.c, with pitch_est_defines.h's constants (its tables are in
' silk_tables_c).
Option Explicit

' pitch_est_defines.h
'/********************************************************/
'/* Definitions for pitch estimator                      */
'/********************************************************/
Public Const PE_MAX_FS_KHZ As Long = 16                 '/* Maximum sampling frequency used */

Public Const PE_MAX_NB_SUBFR As Long = 4
Public Const PE_SUBFR_LENGTH_MS As Long = 5             '/* 5 ms */

Public Const PE_LTP_MEM_LENGTH_MS As Long = 20          ' ( 4 * PE_SUBFR_LENGTH_MS )

Public Const PE_MAX_FRAME_LENGTH_MS As Long = 40        ' ( PE_LTP_MEM_LENGTH_MS + PE_MAX_NB_SUBFR * PE_SUBFR_LENGTH_MS )
Public Const PE_MAX_FRAME_LENGTH As Long = 640          ' ( PE_MAX_FRAME_LENGTH_MS * PE_MAX_FS_KHZ )
Public Const PE_MAX_FRAME_LENGTH_ST_1 As Long = 160     ' ( PE_MAX_FRAME_LENGTH >> 2 )
Public Const PE_MAX_FRAME_LENGTH_ST_2 As Long = 320     ' ( PE_MAX_FRAME_LENGTH >> 1 )

Public Const PE_MAX_LAG_MS As Long = 18                 '/* 18 ms -> 56 Hz */
Public Const PE_MIN_LAG_MS As Long = 2                  '/* 2 ms -> 500 Hz */
Public Const PE_MAX_LAG As Long = 288                   ' ( PE_MAX_LAG_MS * PE_MAX_FS_KHZ )
Public Const PE_MIN_LAG As Long = 32                    ' ( PE_MIN_LAG_MS * PE_MAX_FS_KHZ )

Public Const PE_D_SRCH_LENGTH As Long = 24

Public Const PE_NB_STAGE3_LAGS As Long = 5

Public Const PE_NB_CBKS_STAGE2 As Long = 3
Public Const PE_NB_CBKS_STAGE2_EXT As Long = 11

Public Const PE_NB_CBKS_STAGE3_MAX As Long = 34
Public Const PE_NB_CBKS_STAGE3_MID As Long = 24
Public Const PE_NB_CBKS_STAGE3_MIN As Long = 16

Public Const PE_NB_CBKS_STAGE3_10MS As Long = 12
Public Const PE_NB_CBKS_STAGE2_10MS As Long = 3

Public Const PE_SHORTLAG_BIAS As Double = 0.2           '/* for logarithmic weighting    */
Public Const PE_PREVLAG_BIAS As Double = 0.2            '/* for logarithmic weighting    */
Public Const PE_FLATCONTOUR_BIAS As Double = 0.05

Public Const SILK_PE_MIN_COMPLEX As Long = 0
Public Const SILK_PE_MID_COMPLEX As Long = 1
Public Const SILK_PE_MAX_COMPLEX As Long = 2

'/***********************************************************
'* Pitch analyser function
'********************************************************** */
Public Sub silk_decode_pitch(ByVal lagIndex As Long, ByVal contourIndex As Long, pitch_lags() As Long, _
        ByVal Fs_kHz As Long, ByVal nb_subfr As Long)
    Dim lag As Long, k As Long, min_lag As Long, max_lag As Long, cbk_size As Long
    Dim Lag_CB_ptr As Long                              ' const opus_int8 *: an offset into silk_rom_i8

    If Fs_kHz = 8 Then
        If nb_subfr = PE_MAX_NB_SUBFR Then
            Lag_CB_ptr = silk_CB_lags_stage2
            cbk_size = PE_NB_CBKS_STAGE2_EXT
        Else
            ' celt_assert( nb_subfr == PE_MAX_NB_SUBFR >> 1 );
            Lag_CB_ptr = silk_CB_lags_stage2_10_ms
            cbk_size = PE_NB_CBKS_STAGE2_10MS
        End If
    Else
        If nb_subfr = PE_MAX_NB_SUBFR Then
            Lag_CB_ptr = silk_CB_lags_stage3
            cbk_size = PE_NB_CBKS_STAGE3_MAX
        Else
            ' celt_assert( nb_subfr == PE_MAX_NB_SUBFR >> 1 );
            Lag_CB_ptr = silk_CB_lags_stage3_10_ms
            cbk_size = PE_NB_CBKS_STAGE3_10MS
        End If
    End If

    min_lag = silk_SMULBB(PE_MIN_LAG_MS, Fs_kHz)
    max_lag = silk_SMULBB(PE_MAX_LAG_MS, Fs_kHz)
    lag = min_lag + lagIndex

    For k = 0 To nb_subfr - 1
        pitch_lags(k) = lag + silk_rom_i8(Lag_CB_ptr + k * cbk_size + contourIndex)     ' matrix_ptr( Lag_CB_ptr, k, contourIndex, cbk_size )
        pitch_lags(k) = silk_LIMIT(pitch_lags(k), min_lag, max_lag)
    Next
End Sub
