Attribute VB_Name = "celt_pitch_c"
' ref/opus/celt (libopus 1.5.2): pitch.c and pitch.h, the float build, C paths (no OVERRIDE_*: arch 0)
' (cluster K, fork O2). pitch.h's inlines xcorr_kernel_c, dual_inner_prod_c and celt_inner_prod_c are here as
' functions; pitch.c's find_best_pitch, celt_fir5, pitch_downsample, celt_pitch_xcorr_c (the unrolled
' version: the #if 0 simple one is not compiled) and pitch_search are what the decoder's PLC reaches. Not
' ported (the encoder's): compute_pitch_gain, remove_doubling (and its second_check table).
' (VB) An opus_val16 * / opus_val32 * is a Single array and an index; celt_sig *x[] (pitch_downsample) is the
' decoder's memory and the channels' offsets in it. opus_val32 sum[4] is four ByRef Singles.
Option Explicit

' (VB) pitch_search's x_lp4, y_lp4 and xcorr (ALLOC'd in the C)
Private x_lp4() As Single
Private y_lp4() As Single
Private xcorr_buf() As Single
Private pitch_ac(0 To 4) As Single
Private pitch_done As Boolean

Private Sub pitch_init()
    ReDim x_lp4(0 To 1023)
    ReDim y_lp4(0 To 1023)
    ReDim xcorr_buf(0 To 1023)
    pitch_done = True
End Sub

' OPT: This is the kernel you really want to optimize. It gets used a lot
'   by the prefilter and by the PLC.
Public Sub xcorr_kernel_c(x() As Single, ByVal xo As Long, y() As Single, ByVal yo As Long, sum0 As Single, sum1 As Single, sum2 As Single, sum3 As Single, ByVal len_ As Long)
    Dim j As Long, y_0 As Single, y_1 As Single, y_2 As Single, y_3 As Single, tmp As Single
    ' celt_assert(len>=3);
    y_3 = 0                                             ' gcc doesn't realize that y_3 can't be used uninitialized
    y_0 = y(yo): yo = yo + 1
    y_1 = y(yo): yo = yo + 1
    y_2 = y(yo): yo = yo + 1
    j = 0
    Do While j < len_ - 3
        tmp = x(xo): xo = xo + 1
        y_3 = y(yo): yo = yo + 1
        sum0 = sum0 + tmp * y_0
        sum1 = sum1 + tmp * y_1
        sum2 = sum2 + tmp * y_2
        sum3 = sum3 + tmp * y_3
        tmp = x(xo): xo = xo + 1
        y_0 = y(yo): yo = yo + 1
        sum0 = sum0 + tmp * y_1
        sum1 = sum1 + tmp * y_2
        sum2 = sum2 + tmp * y_3
        sum3 = sum3 + tmp * y_0
        tmp = x(xo): xo = xo + 1
        y_1 = y(yo): yo = yo + 1
        sum0 = sum0 + tmp * y_2
        sum1 = sum1 + tmp * y_3
        sum2 = sum2 + tmp * y_0
        sum3 = sum3 + tmp * y_1
        tmp = x(xo): xo = xo + 1
        y_2 = y(yo): yo = yo + 1
        sum0 = sum0 + tmp * y_3
        sum1 = sum1 + tmp * y_0
        sum2 = sum2 + tmp * y_1
        sum3 = sum3 + tmp * y_2
        j = j + 4
    Loop
    j = j + 1
    If j - 1 < len_ Then                                ' if (j++<len)
        tmp = x(xo): xo = xo + 1
        y_3 = y(yo): yo = yo + 1
        sum0 = sum0 + tmp * y_0
        sum1 = sum1 + tmp * y_1
        sum2 = sum2 + tmp * y_2
        sum3 = sum3 + tmp * y_3
    End If
    j = j + 1
    If j - 1 < len_ Then                                ' if (j++<len)
        tmp = x(xo): xo = xo + 1
        y_0 = y(yo): yo = yo + 1
        sum0 = sum0 + tmp * y_1
        sum1 = sum1 + tmp * y_2
        sum2 = sum2 + tmp * y_3
        sum3 = sum3 + tmp * y_0
    End If
    If j < len_ Then
        tmp = x(xo): xo = xo + 1
        y_1 = y(yo): yo = yo + 1
        sum0 = sum0 + tmp * y_2
        sum1 = sum1 + tmp * y_3
        sum2 = sum2 + tmp * y_0
        sum3 = sum3 + tmp * y_1
    End If
End Sub

' #ifndef OVERRIDE_DUAL_INNER_PROD: dual_inner_prod(x, y01, y02, N, xy1, xy2, arch) is dual_inner_prod_c
Public Sub dual_inner_prod_c(x() As Single, ByVal xo As Long, y01() As Single, ByVal y01o As Long, y02() As Single, ByVal y02o As Long, ByVal N As Long, xy1 As Single, xy2 As Single)
    Dim i As Long, xy01 As Single, xy02 As Single
    xy01 = 0
    xy02 = 0
    For i = 0 To N - 1
        xy01 = xy01 + x(xo + i) * y01(y01o + i)         ' MAC16_16(xy01, x[i], y01[i])
        xy02 = xy02 + x(xo + i) * y02(y02o + i)
    Next
    xy1 = xy01
    xy2 = xy02
End Sub

'We make sure a C version is always available for cases where the overhead of
'  vectorization and passing around an arch flag aren't worth it.
' (#if !defined(OVERRIDE_CELT_INNER_PROD): celt_inner_prod(x, y, N, arch) is celt_inner_prod_c)
Public Function celt_inner_prod_c(x() As Single, ByVal xo As Long, y() As Single, ByVal yo As Long, ByVal N As Long) As Single
    Dim i As Long, xy As Single
    xy = 0
    For i = 0 To N - 1
        xy = xy + x(xo + i) * y(yo + i)                 ' MAC16_16(xy, x[i], y[i])
    Next
    celt_inner_prod_c = xy
End Function

' pitch.c

Private Sub find_best_pitch(xcorr() As Single, y() As Single, ByVal yo As Long, ByVal len_ As Long, ByVal max_pitch As Long, best_pitch() As Long)
    Dim i As Long, j As Long, Syy As Single, best_num(0 To 1) As Single, best_den(0 To 1) As Single
    Dim num As Single, xcorr16 As Single
    ' #ifdef FIXED_POINT ... #endif
    Syy = 1
    best_num(0) = -1
    best_num(1) = -1
    best_den(0) = 0
    best_den(1) = 0
    best_pitch(0) = 0
    best_pitch(1) = 1
    For j = 0 To len_ - 1
        Syy = Syy + y(yo + j) * y(yo + j)               ' ADD32(Syy, SHR32(MULT16_16(y[j],y[j]), yshift))
    Next
    For i = 0 To max_pitch - 1
        If xcorr(i) > 0 Then
            xcorr16 = xcorr(i)                          ' EXTRACT16(VSHR32(xcorr[i], xshift))
            ' #ifndef FIXED_POINT
            ' Considering the range of xcorr16, this should avoid both underflows
            '   and overflows (inf) when squaring xcorr16
            xcorr16 = xcorr16 * 1E-12!
            ' #endif
            num = xcorr16 * xcorr16                     ' MULT16_16_Q15(xcorr16,xcorr16)
            If num * best_den(1) > best_num(1) * Syy Then
                If num * best_den(0) > best_num(0) * Syy Then
                    best_num(1) = best_num(0)
                    best_den(1) = best_den(0)
                    best_pitch(1) = best_pitch(0)
                    best_num(0) = num
                    best_den(0) = Syy
                    best_pitch(0) = i
                Else
                    best_num(1) = num
                    best_den(1) = Syy
                    best_pitch(1) = i
                End If
            End If
        End If
        Syy = Syy + (y(yo + i + len_) * y(yo + i + len_) - y(yo + i) * y(yo + i))
        If Syy < 1 Then Syy = 1                         ' MAX32(1, Syy)
    Next
End Sub

Private Sub celt_fir5(x() As Single, num() As Single, ByVal N As Long)
    Dim i As Long, num0 As Single, num1 As Single, num2 As Single, num3 As Single, num4 As Single
    Dim mem0 As Single, mem1 As Single, mem2 As Single, mem3 As Single, mem4 As Single, sum As Single
    num0 = num(0)
    num1 = num(1)
    num2 = num(2)
    num3 = num(3)
    num4 = num(4)
    mem0 = 0
    mem1 = 0
    mem2 = 0
    mem3 = 0
    mem4 = 0
    For i = 0 To N - 1
        sum = x(i)                                      ' SHL32(EXTEND32(x[i]), SIG_SHIFT)
        sum = sum + num0 * mem0
        sum = sum + num1 * mem1
        sum = sum + num2 * mem2
        sum = sum + num3 * mem3
        sum = sum + num4 * mem4
        mem4 = mem3
        mem3 = mem2
        mem2 = mem1
        mem1 = mem0
        mem0 = x(i)
        x(i) = sum                                      ' ROUND16(sum, SIG_SHIFT)
    Next
End Sub

' (VB) celt_sig *x[]: x() with the channels at offsets x0o and x1o; x_lp is x_lp() from 0
Public Sub pitch_downsample(x() As Single, ByVal x0o As Long, ByVal x1o As Long, x_lp() As Single, ByVal len_ As Long, ByVal C As Long, ByVal arch As Long)
    Dim i As Long, tmp As Single, lpc(0 To 3) As Single, lpc2(0 To 4) As Single, c1 As Single, t As Single
    tmp = 1!                                            ' Q15ONE
    c1 = 0.8!                                           ' QCONST16(.8f,15)
    ' #ifdef FIXED_POINT ... #else
    For i = 1 To (len_ \ 2) - 1
        x_lp(i) = 0.25! * x(x0o + (2 * i - 1)) + 0.25! * x(x0o + (2 * i + 1)) + 0.5! * x(x0o + 2 * i)
    Next
    x_lp(0) = 0.25! * x(x0o + 1) + 0.5! * x(x0o)
    If C = 2 Then
        For i = 1 To (len_ \ 2) - 1
            x_lp(i) = x_lp(i) + (0.25! * x(x1o + (2 * i - 1)) + 0.25! * x(x1o + (2 * i + 1)) + 0.5! * x(x1o + 2 * i))
        Next
        x_lp(0) = x_lp(0) + (0.25! * x(x1o + 1) + 0.5! * x(x1o))
    End If
    ' #endif
    celt_autocorr_ x_lp, 0, pitch_ac, x_lp, 0, 4, len_ \ 2, arch    ' (window NULL: not read with overlap 0)

    ' Noise floor -40 dB
    ' #ifdef FIXED_POINT ... #else
    pitch_ac(0) = pitch_ac(0) * 1.0001!
    ' #endif
    ' Lag windowing
    For i = 1 To 4
        'ac[i] *= exp(-.5*(2*M_PI*.002*i)*(2*M_PI*.002*i));
        ' #ifdef FIXED_POINT ... #else
        t = 0.008! * i
        pitch_ac(i) = pitch_ac(i) - pitch_ac(i) * t * t ' ac[i] -= ac[i]*(.008f*i)*(.008f*i);
    Next

    celt_lpc_ lpc, 0, pitch_ac, 4
    For i = 0 To 3
        tmp = 0.9! * tmp                                ' MULT16_16_Q15(QCONST16(.9f,15), tmp)
        lpc(i) = lpc(i) * tmp
    Next
    ' Add a zero
    lpc2(0) = lpc(0) + 0.8!                             ' QCONST16(.8f,SIG_SHIFT)
    lpc2(1) = lpc(1) + c1 * lpc(0)
    lpc2(2) = lpc(2) + c1 * lpc(1)
    lpc2(3) = lpc(3) + c1 * lpc(2)
    lpc2(4) = c1 * lpc(3)
    celt_fir5 x_lp, lpc2, len_ \ 2
End Sub

' Pure C implementation.
' (VB) the float build's: no return value
Public Sub celt_pitch_xcorr_c(x_() As Single, ByVal xo As Long, y_() As Single, ByVal yo As Long, xcorr() As Single, ByVal len_ As Long, ByVal max_pitch As Long, ByVal arch As Long)
    Dim i As Long, sum0 As Single, sum1 As Single, sum2 As Single, sum3 As Single
    ' #if 0: the simple version (not compiled)
    ' #else: Unrolled version of the pitch correlation -- runs faster on x86 and ARM
    'The EDSP version requires that max_pitch is at least 1, and that _x is
    '   32-bit aligned.
    '  Since it's hard to put asserts in assembly, put them here.
    ' celt_assert(max_pitch>0);
    ' celt_sig_assert(((size_t)_x&3)==0);
    For i = 0 To max_pitch - 4 Step 4                   ' for (i=0;i<max_pitch-3;i+=4)
        sum0 = 0: sum1 = 0: sum2 = 0: sum3 = 0
        xcorr_kernel_c x_, xo, y_, yo + i, sum0, sum1, sum2, sum3, len_
        xcorr(i) = sum0
        xcorr(i + 1) = sum1
        xcorr(i + 2) = sum2
        xcorr(i + 3) = sum3
    Next
    ' In case max_pitch isn't a multiple of 4, do non-unrolled version.
    For i = i To max_pitch - 1
        xcorr(i) = celt_inner_prod_c(x_, xo, y_, yo + i, len_)
    Next
End Sub

' (VB) x_lp: x_lp() from xo; y: y() from yo
Public Sub pitch_search(x_lp() As Single, ByVal xo As Long, y() As Single, ByVal yo As Long, ByVal len_ As Long, ByVal max_pitch As Long, pitch As Long, ByVal arch As Long)
    Dim i As Long, j As Long, lag As Long, best_pitch(0 To 1) As Long, offset As Long, sum As Single
    Dim a As Single, b As Single, c As Single
    If Not pitch_done Then pitch_init

    ' celt_assert(len>0);
    ' celt_assert(max_pitch>0);
    lag = len_ + max_pitch

    ' ALLOC(x_lp4, len>>2, opus_val16); ALLOC(y_lp4, lag>>2, opus_val16); ALLOC(xcorr, max_pitch>>1, opus_val32);

    ' Downsample by 2 again
    For j = 0 To (len_ \ 4) - 1
        x_lp4(j) = x_lp(xo + 2 * j)
    Next
    For j = 0 To (lag \ 4) - 1
        y_lp4(j) = y(yo + 2 * j)
    Next

    ' #ifdef FIXED_POINT ... #endif

    ' Coarse search with 4x decimation

    celt_pitch_xcorr_c x_lp4, 0, y_lp4, 0, xcorr_buf, len_ \ 4, max_pitch \ 4, arch

    find_best_pitch xcorr_buf, y_lp4, 0, len_ \ 4, max_pitch \ 4, best_pitch

    ' Finer search with 2x decimation
    For i = 0 To (max_pitch \ 2) - 1
        xcorr_buf(i) = 0
        If Abs(i - 2 * best_pitch(0)) > 2 And Abs(i - 2 * best_pitch(1)) > 2 Then GoTo continue_
        ' #ifdef FIXED_POINT ... #else
        sum = celt_inner_prod_c(x_lp, xo, y, yo + i, len_ \ 2)
        ' #endif
        If sum < -1 Then xcorr_buf(i) = -1 Else xcorr_buf(i) = sum     ' MAX32(-1, sum)
continue_:
    Next
    find_best_pitch xcorr_buf, y, yo, len_ \ 2, max_pitch \ 2, best_pitch

    ' Refine by pseudo-interpolation
    If best_pitch(0) > 0 And best_pitch(0) < (max_pitch \ 2) - 1 Then
        a = xcorr_buf(best_pitch(0) - 1)
        b = xcorr_buf(best_pitch(0))
        c = xcorr_buf(best_pitch(0) + 1)
        If (c - a) > 0.7! * (b - a) Then                ' MULT16_32_Q15(QCONST16(.7f,15),b-a)
            offset = 1
        ElseIf (a - c) > 0.7! * (b - c) Then
            offset = -1
        Else
            offset = 0
        End If
    Else
        offset = 0
    End If
    pitch = 2 * best_pitch(0) - offset
End Sub
