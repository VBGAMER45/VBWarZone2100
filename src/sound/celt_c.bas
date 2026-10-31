Attribute VB_Name = "celt_c"
' ref/opus/celt (libopus 1.5.2): celt.c and celt.h, the float build without CUSTOM_MODES, C paths (no
' OPUS_ARM_ASM, no OVERRIDE_comb_filter) (cluster K, fork O2): resampling_factor, comb_filter_const_c,
' comb_filter, tf_select_table, init_caps, and celt.h's request constants and icdf tables. opus_strerror and
' opus_get_version_string (also in celt.c) are O3's (opus_c.bas, K_DESIGN.md's contract). celt_preemphasis
' (declared in celt.h) is the encoder's. celt.h's CUSTOM_MODES toOpus / fromOpus are not compiled.
' (VB) An opus_val32 * is a Single array and an index (negative offsets from it are read, as the C does).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' celt.h

Public Const LEAK_BANDS As Long = 19

' Encoder/decoder Requests

Public Const CELT_SET_PREDICTION_REQUEST As Long = 10002
'* Controls the use of interframe prediction.
'    0=Independent frames
'    1=Short term interframe prediction allowed
'    2=Long term prediction allowed

Public Const CELT_SET_INPUT_CLIPPING_REQUEST As Long = 10004

Public Const CELT_GET_AND_CLEAR_ERROR_REQUEST As Long = 10007

Public Const CELT_SET_CHANNELS_REQUEST As Long = 10008

' Internal
Public Const CELT_SET_START_BAND_REQUEST As Long = 10010

Public Const CELT_SET_END_BAND_REQUEST As Long = 10012

Public Const CELT_GET_MODE_REQUEST As Long = 10015
'* Get the CELTMode used by an encoder or decoder

Public Const CELT_SET_SIGNALLING_REQUEST As Long = 10016

Public Const CELT_SET_TONALITY_REQUEST As Long = 10018
Public Const CELT_SET_TONALITY_SLOPE_REQUEST As Long = 10020

Public Const CELT_SET_ANALYSIS_REQUEST As Long = 10022

Public Const OPUS_SET_LFE_REQUEST As Long = 10024

Public Const OPUS_SET_ENERGY_MASK_REQUEST As Long = 10026

Public Const CELT_SET_SILK_INFO_REQUEST As Long = 10028

Public Const COMBFILTER_MAXPERIOD As Long = 1024
Public Const COMBFILTER_MINPERIOD As Long = 15

' static const unsigned char trim_icdf[11] = {126, 124, 119, 109, 87, 41, 19, 9, 4, 2, 0};
' Probs: NONE: 21.875%, LIGHT: 6.25%, NORMAL: 65.625%, AGGRESSIVE: 6.25%
' static const unsigned char spread_icdf[4] = {25, 23, 2, 0};
' static const unsigned char tapset_icdf[3]={2,1,0};
Public trim_icdf() As Byte
Public spread_icdf() As Byte
Public tapset_icdf() As Byte

' TF change table. Positive values mean better frequency resolution (longer
'   effective window), whereas negative values mean better time resolution
'   (shorter effective window). The second index is computed as:
'   4*isTransient + 2*tf_select + per_band_flag
' const signed char tf_select_table[4][8]: (VB) flattened, [LM][k] is tf_select_table(LM * 8 + k)
Public tf_select_table() As Long

' comb_filter's static const opus_val16 gains[3][3]: (VB) [t][k] is comb_gains(t * 3 + k)
Private comb_gains(0 To 8) As Single

Private celt_c_done As Boolean

' (VB) the static tables (called once by the decoder's init)
Public Sub celt_c_init()
    Dim v As Variant, i As Long
    If celt_c_done Then Exit Sub
    ReDim trim_icdf(0 To 10)
    v = Array(126, 124, 119, 109, 87, 41, 19, 9, 4, 2, 0)
    For i = 0 To 10
        trim_icdf(i) = v(i)
    Next
    ReDim spread_icdf(0 To 3)
    spread_icdf(0) = 25: spread_icdf(1) = 23: spread_icdf(2) = 2: spread_icdf(3) = 0
    ReDim tapset_icdf(0 To 2)
    tapset_icdf(0) = 2: tapset_icdf(1) = 1: tapset_icdf(2) = 0
    ReDim tf_select_table(0 To 31)
    '                 isTransient=0     isTransient=1
    v = Array(0, -1, 0, -1, 0, -1, 0, -1, _
              0, -1, 0, -2, 1, 0, 1, -1, _
              0, -2, 0, -3, 2, 0, 1, -1, _
              0, -2, 0, -3, 3, 0, 1, -1)
    ' 2.5 ms, 5 ms, 10 ms, 20 ms
    For i = 0 To 31
        tf_select_table(i) = v(i)
    Next
    ' {QCONST16(0.3066406250f, 15), QCONST16(0.2170410156f, 15), QCONST16(0.1296386719f, 15)},
    ' {QCONST16(0.4638671875f, 15), QCONST16(0.2680664062f, 15), QCONST16(0.f, 15)},
    ' {QCONST16(0.7998046875f, 15), QCONST16(0.1000976562f, 15), QCONST16(0.f, 15)}
    comb_gains(0) = 0.306640625!: comb_gains(1) = 0.2170410156!: comb_gains(2) = 0.1296386719!
    comb_gains(3) = 0.4638671875!: comb_gains(4) = 0.2680664062!: comb_gains(5) = 0!
    comb_gains(6) = 0.7998046875!: comb_gains(7) = 0.1000976562!: comb_gains(8) = 0!
    celt_c_done = True
End Sub

' celt.c

Public Function resampling_factor(ByVal rate As Long) As Long
    Dim ret As Long
    Select Case rate
    Case 48000
        ret = 1
    Case 24000
        ret = 2
    Case 16000
        ret = 3
    Case 12000
        ret = 4
    Case 8000
        ret = 6
    Case Else
        ' #ifndef CUSTOM_MODES: celt_assert(0); (compiled out)
        ret = 0
    End Select
    resampling_factor = ret
End Function

' #if !defined(OVERRIDE_COMB_FILTER_CONST) || defined(NON_STATIC_COMB_FILTER_CONST_C)
' (#ifdef OPUS_ARM_ASM: the ARM version, not compiled)
' (VB) y: y() from yo; x: x() from xo
Private Sub comb_filter_const_c(y() As Single, ByVal yo As Long, x() As Single, ByVal xo As Long, ByVal T As Long, ByVal N As Long, ByVal g10 As Single, ByVal g11 As Single, ByVal g12 As Single)
    Dim x0 As Single, x1 As Single, x2 As Single, x3 As Single, x4 As Single, i As Long, k As Long
    x4 = x(xo - T - 2)
    x3 = x(xo - T - 1)
    x2 = x(xo - T)
    x1 = x(xo - T + 1)
    k = xo - T + 2
    For i = 0 To N - 1
        x0 = x(k + i)                                   ' x[i-T+2]
        y(yo + i) = x(xo + i) + g10 * x2 + g11 * (x1 + x3) + g12 * (x0 + x4)
        ' y[i] = SATURATE(y[i], SIG_SAT);  (float build: nothing)
        x4 = x3
        x3 = x2
        x2 = x1
        x1 = x0
    Next
End Sub

' #ifndef OVERRIDE_comb_filter
' (VB) y: y() from yo; x: x() from xo; window: window() from 0 (not read when overlap is 0)
Public Sub comb_filter(y() As Single, ByVal yo As Long, x() As Single, ByVal xo As Long, ByVal T0 As Long, ByVal T1 As Long, ByVal N As Long, _
        ByVal g0 As Single, ByVal g1 As Single, ByVal tapset0 As Long, ByVal tapset1 As Long, window() As Single, ByVal overlap As Long, ByVal arch As Long)
    Dim i As Long, g00 As Single, g01 As Single, g02 As Single, g10 As Single, g11 As Single, g12 As Single
    Dim x0 As Single, x1 As Single, x2 As Single, x3 As Single, x4 As Single, f As Single, nf As Single, k As Long
    ' printf ("%d %d %f %f\n", T0, T1, g0, g1);

    If g0 = 0 And g1 = 0 Then
        ' OPT: Happens to work without the OPUS_MOVE(), but only because the current encoder already copies x to y
        If VarPtr(x(xo)) <> VarPtr(y(yo)) Then CopyMemory y(yo), x(xo), N * 4     ' OPUS_MOVE(y, x, N)
        Exit Sub
    End If
    ' When the gain is zero, T0 and/or T1 is set to zero. We need
    '   to have then be at least 2 to avoid processing garbage data.
    If T0 < COMBFILTER_MINPERIOD Then T0 = COMBFILTER_MINPERIOD
    If T1 < COMBFILTER_MINPERIOD Then T1 = COMBFILTER_MINPERIOD
    g00 = g0 * comb_gains(tapset0 * 3)                  ' MULT16_16_P15(g0, gains[tapset0][0])
    g01 = g0 * comb_gains(tapset0 * 3 + 1)
    g02 = g0 * comb_gains(tapset0 * 3 + 2)
    g10 = g1 * comb_gains(tapset1 * 3)
    g11 = g1 * comb_gains(tapset1 * 3 + 1)
    g12 = g1 * comb_gains(tapset1 * 3 + 2)
    x1 = x(xo - T1 + 1)
    x2 = x(xo - T1)
    x3 = x(xo - T1 - 1)
    x4 = x(xo - T1 - 2)
    ' If the filter didn't change, we don't need the overlap
    If g0 = g1 And T0 = T1 And tapset0 = tapset1 Then overlap = 0
    For i = 0 To overlap - 1
        x0 = x(xo + i - T1 + 2)
        f = window(i) * window(i)                       ' MULT16_16_Q15(window[i],window[i])
        nf = 1! - f                                     ' (Q15ONE-f)
        k = xo + i - T0
        y(yo + i) = x(xo + i) _
                 + (nf * g00) * x(k) _
                 + (nf * g01) * (x(k + 1) + x(k - 1)) _
                 + (nf * g02) * (x(k + 2) + x(k - 2)) _
                 + (f * g10) * x2 _
                 + (f * g11) * (x1 + x3) _
                 + (f * g12) * (x0 + x4)
        ' y[i] = SATURATE(y[i], SIG_SAT);  (float build: nothing)
        x4 = x3
        x3 = x2
        x2 = x1
        x1 = x0
    Next
    If g1 = 0 Then
        ' OPT: Happens to work without the OPUS_MOVE(), but only because the current encoder already copies x to y
        If VarPtr(x(xo)) <> VarPtr(y(yo)) Then CopyMemory y(yo + overlap), x(xo + overlap), (N - overlap) * 4
        Exit Sub
    End If

    ' Compute the part with the constant filter.
    comb_filter_const_c y, yo + i, x, xo + i, T1, N - i, g10, g11, g12  ' comb_filter_const(y+i, x+i, T1, N-i, g10, g11, g12, arch)
End Sub
' #endif /* OVERRIDE_comb_filter */

' (VB) int *cap is cap() from 0
Public Sub init_caps(ByVal m As Long, cap() As Long, ByVal LM As Long, ByVal C As Long)
    Dim i As Long, N As Long
    With celt_modes(m)
        For i = 0 To .nbEBands - 1
            N = (.eBands(i + 1) - .eBands(i)) * celt_pow2(LM)
            cap(i) = ((CLng(.cache.caps(.nbEBands * (2 * LM + C - 1) + i)) + 64) * C * N) \ 4
        Next
    End With
End Sub
