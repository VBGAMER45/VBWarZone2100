Attribute VB_Name = "celt_mdct_c"
' ref/opus/celt (libopus 1.5.2): mdct.c, the float build without CUSTOM_MODES (cluster K, fork O2): the
' backward MDCT the decoder uses (clt_mdct_backward_c; no OVERRIDE_clt_mdct_backward: arch 0, C path). Not
' ported: clt_mdct_forward_c (the encoder's), and clt_mdct_init / clt_mdct_clear (CUSTOM_MODES only: the
' static mode's mdct_lookup is celt_modes_c's).
'
' This is a simple MDCT implementation that uses a N/4 complex FFT
'   to do most of the work. It should be relatively straightforward to
'   plug in pretty much and FFT here.
'
'   This replaces the Vorbis FFT (and uses the exact same API), which
'   was a bit too messy and that was ending up duplicating code
'   (might as well use the same FFT everywhere).
'
'   The algorithm is similar to (and inspired from) Fabrice Bellard's
'   MDCT implementation in FFMPEG, but has differences in signs, ordering
'   and scaling in many places.
' (VB) A kiss_fft_scalar * is a Single array and an index.
Option Explicit

' #ifndef OVERRIDE_clt_mdct_backward
' (VB) in: in_() from ino; out: out_() from outo; window: window() from 0
Public Sub clt_mdct_backward_c(l As mdct_lookup, in_() As Single, ByVal ino As Long, out_() As Single, ByVal outo As Long, window() As Single, ByVal overlap As Long, ByVal shift As Long, ByVal stride As Long, ByVal arch As Long)
    Dim i As Long, N As Long, N2 As Long, N4 As Long, trig As Long
    Dim xp1 As Long, xp2 As Long, yp As Long, rev As Long, yr As Single, yi As Single
    Dim yp0 As Long, yp1 As Long, re As Single, im As Single, t0 As Single, t1 As Single
    Dim wp1 As Long, wp2 As Long, x1 As Single, x2 As Single, kf As Long, stride2 As Long
    ' (void) arch;

    N = l.n
    trig = 0                                            ' trig = l->trig;
    For i = 0 To shift - 1
        N = N \ 2
        trig = trig + N
    Next
    N2 = N \ 2
    N4 = N \ 4
    kf = l.kfft(shift)

    ' Pre-rotate
    ' Temp pointers to make it really clear to the compiler what we're doing
    xp1 = ino
    xp2 = ino + stride * (N2 - 1)
    yp = outo + overlap \ 2
    stride2 = 2 * stride
    ' const kiss_twiddle_scalar * OPUS_RESTRICT t = &trig[0];
    ' const opus_int16 * OPUS_RESTRICT bitrev = l->kfft[shift]->bitrev;
    For i = 0 To N4 - 1
        rev = celt_fft_states(kf).bitrev(i)
        yr = in_(xp2) * l.trig(trig + i) + in_(xp1) * l.trig(trig + N4 + i)
        yi = in_(xp1) * l.trig(trig + i) - in_(xp2) * l.trig(trig + N4 + i)
        ' We swap real and imag because we use an FFT instead of an IFFT.
        out_(yp + 2 * rev + 1) = yr
        out_(yp + 2 * rev) = yi
        ' Storing the pre-rotation directly in the bitrev order.
        xp1 = xp1 + stride2
        xp2 = xp2 - stride2
    Next

    opus_fft_impl celt_fft_states(kf), out_, outo + overlap \ 2

    ' Post-rotate and de-shuffle from both ends of the buffer at once to make
    '   it in-place.
    yp0 = outo + overlap \ 2
    yp1 = outo + overlap \ 2 + N2 - 2
    ' Loop to (N4+1)>>1 to handle odd N4. When N4 is odd, the
    '   middle pair will be computed twice.
    For i = 0 To ((N4 + 1) \ 2) - 1
        ' We swap real and imag because we're using an FFT instead of an IFFT.
        re = out_(yp0 + 1)
        im = out_(yp0)
        t0 = l.trig(trig + i)
        t1 = l.trig(trig + N4 + i)
        ' We'd scale up by 2 here, but instead it's done when mixing the windows
        yr = re * t0 + im * t1
        yi = re * t1 - im * t0
        ' We swap real and imag because we're using an FFT instead of an IFFT.
        re = out_(yp1 + 1)
        im = out_(yp1)
        out_(yp0) = yr
        out_(yp1 + 1) = yi

        t0 = l.trig(trig + (N4 - i - 1))
        t1 = l.trig(trig + (N2 - i - 1))
        ' We'd scale up by 2 here, but instead it's done when mixing the windows
        yr = re * t0 + im * t1
        yi = re * t1 - im * t0
        out_(yp1) = yr
        out_(yp0 + 1) = yi
        yp0 = yp0 + 2
        yp1 = yp1 - 2
    Next

    ' Mirror on both sides for TDAC
    xp1 = outo + overlap - 1
    yp1 = outo
    wp1 = 0
    wp2 = overlap - 1
    For i = 0 To overlap \ 2 - 1
        x1 = out_(xp1)
        x2 = out_(yp1)
        out_(yp1) = window(wp2) * x2 - window(wp1) * x1
        yp1 = yp1 + 1
        out_(xp1) = window(wp1) * x2 + window(wp2) * x1
        xp1 = xp1 - 1
        wp1 = wp1 + 1
        wp2 = wp2 - 1
    Next
End Sub
' #endif /* OVERRIDE_clt_mdct_backward */
