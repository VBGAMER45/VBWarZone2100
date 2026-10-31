Attribute VB_Name = "celt_lpc_c"
' ref/opus/celt (libopus 1.5.2): celt_lpc.c and celt_lpc.h, the float build, C paths (no OVERRIDE_CELT_FIR,
' SMALL_FOOTPRINT undefined: the unrolled celt_iir) (cluster K, fork O2): _celt_lpc, celt_fir_c, celt_iir,
' _celt_autocorr, which the decoder's PLC reaches. The FIXED_POINT parts are not compiled in this build.
' (VB) Names lose their leading underscore at the front and get one at the end (_celt_lpc is celt_lpc_,
' _celt_autocorr celt_autocorr_). An opus_val16 * / opus_val32 * is a Single array and an index.
Option Explicit

Public Const CELT_LPC_ORDER As Long = 24

' (VB) the C's ALLOC'd scratch: celt_fir_c's rnum, celt_iir's rden and y, _celt_autocorr's xx
Private rnum() As Single
Private rden() As Single
Private iir_y() As Single
Private autocorr_xx() As Single
Private lpc_done As Boolean

Private Sub lpc_init()
    ReDim rnum(0 To CELT_LPC_ORDER - 1)
    ReDim rden(0 To CELT_LPC_ORDER - 1)
    ReDim iir_y(0 To 2047)
    ReDim autocorr_xx(0 To 2047)
    lpc_done = True
End Sub

' (VB) _lpc: lpc_() from lpco (out: [0...p-1] LPC coefficients); ac: ac() from 0 (in: [0...p] autocorrelation values)
Public Sub celt_lpc_(lpc_() As Single, ByVal lpco As Long, ac() As Single, ByVal p As Long)
    Dim i As Long, j As Long, r As Single, error_ As Single, rr As Single, tmp1 As Single, tmp2 As Single
    error_ = ac(0)
    ' #ifdef FIXED_POINT ... #else: float *lpc = _lpc;

    For i = 0 To p - 1                                  ' OPUS_CLEAR(lpc, p);
        lpc_(lpco + i) = 0
    Next
    ' #ifdef FIXED_POINT ... #else
    If ac(0) > 1E-10! Then
        For i = 0 To p - 1
            ' Sum up this iteration's reflection coefficient
            rr = 0
            For j = 0 To i - 1
                rr = rr + lpc_(lpco + j) * ac(i - j)    ' MULT32_32_Q31(lpc[j],ac[i - j])
            Next
            rr = rr + ac(i + 1)                         ' SHR32(ac[i + 1],6)
            r = -(rr / error_)                          ' -frac_div32(SHL32(rr,6), error)
            '  Update LPC coefficients and total error
            lpc_(lpco + i) = r                          ' SHR32(r,6)
            For j = 0 To ((i + 1) \ 2) - 1
                tmp1 = lpc_(lpco + j)
                tmp2 = lpc_(lpco + i - 1 - j)
                lpc_(lpco + j) = tmp1 + r * tmp2        ' MULT32_32_Q31(r,tmp2)
                lpc_(lpco + i - 1 - j) = tmp2 + r * tmp1
            Next

            error_ = error_ - (r * r) * error_          ' MULT32_32_Q31(MULT32_32_Q31(r,r),error)
            ' Bail out once we get 30 dB gain
            ' #ifdef FIXED_POINT ... #else
            If error_ <= 0.001! * ac(0) Then Exit For
        Next
    End If
    ' #ifdef FIXED_POINT: the conversion to int16 (not compiled)
End Sub

' (VB) x: x() from xo; num: num() from numo; y: y() from yo
Public Sub celt_fir_c(x() As Single, ByVal xo As Long, num() As Single, ByVal numo As Long, y() As Single, ByVal yo As Long, ByVal N As Long, ByVal ord As Long, ByVal arch As Long)
    Dim i As Long, j As Long, sum0 As Single, sum1 As Single, sum2 As Single, sum3 As Single, sum As Single
    If Not lpc_done Then lpc_init
    ' celt_assert(x != y);
    ' ALLOC(rnum, ord, opus_val16);
    For i = 0 To ord - 1
        rnum(i) = num(numo + ord - i - 1)
    Next
    For i = 0 To N - 4 Step 4                           ' for (i=0;i<N-3;i+=4)
        sum0 = x(xo + i)                                ' SHL32(EXTEND32(x[i  ]), SIG_SHIFT)
        sum1 = x(xo + i + 1)
        sum2 = x(xo + i + 2)
        sum3 = x(xo + i + 3)
        xcorr_kernel_c rnum, 0, x, xo + i - ord, sum0, sum1, sum2, sum3, ord
        y(yo + i) = sum0                                ' SROUND16(sum[0], SIG_SHIFT)
        y(yo + i + 1) = sum1
        y(yo + i + 2) = sum2
        y(yo + i + 3) = sum3
    Next
    For i = i To N - 1
        sum = x(xo + i)
        For j = 0 To ord - 1
            sum = sum + rnum(j) * x(xo + i + j - ord)   ' MAC16_16(sum,rnum[j],x[i+j-ord])
        Next
        y(yo + i) = sum                                 ' SROUND16(sum, SIG_SHIFT)
    Next
End Sub

' (VB) _x: x_() from xo; den: den() from deno; _y: y_() from yo; mem: mem() from memo
Public Sub celt_iir(x_() As Single, ByVal xo As Long, den() As Single, ByVal deno As Long, y_() As Single, ByVal yo As Long, ByVal N As Long, ByVal ord As Long, mem() As Single, ByVal memo As Long, ByVal arch As Long)
    Dim i As Long, j As Long, sum0 As Single, sum1 As Single, sum2 As Single, sum3 As Single, sum As Single
    Dim den0 As Single, den1 As Single, den2 As Single
    ' #ifdef SMALL_FOOTPRINT ... #else
    If Not lpc_done Then lpc_init
    ' celt_assert((ord&3)==0);
    ' ALLOC(rden, ord, opus_val16); ALLOC(y, N+ord, opus_val16);
    If UBound(iir_y) < N + ord - 1 Then ReDim iir_y(0 To N + ord - 1)
    For i = 0 To ord - 1
        rden(i) = den(deno + ord - i - 1)
    Next
    For i = 0 To ord - 1
        iir_y(i) = -mem(memo + ord - i - 1)
    Next
    For i = ord To N + ord - 1
        iir_y(i) = 0
    Next
    den0 = den(deno)
    den1 = den(deno + 1)
    den2 = den(deno + 2)
    For i = 0 To N - 4 Step 4                           ' for (i=0;i<N-3;i+=4)
        ' Unroll by 4 as if it were an FIR filter
        sum0 = x_(xo + i)
        sum1 = x_(xo + i + 1)
        sum2 = x_(xo + i + 2)
        sum3 = x_(xo + i + 3)
        xcorr_kernel_c rden, 0, iir_y, i, sum0, sum1, sum2, sum3, ord
        ' Patch up the result to compensate for the fact that this is an IIR
        iir_y(i + ord) = -sum0                          ' -SROUND16(sum[0],SIG_SHIFT)
        y_(yo + i) = sum0
        sum1 = sum1 + iir_y(i + ord) * den0             ' MAC16_16(sum[1], y[i+ord  ], den[0])
        iir_y(i + ord + 1) = -sum1
        y_(yo + i + 1) = sum1
        sum2 = sum2 + iir_y(i + ord + 1) * den0
        sum2 = sum2 + iir_y(i + ord) * den1
        iir_y(i + ord + 2) = -sum2
        y_(yo + i + 2) = sum2

        sum3 = sum3 + iir_y(i + ord + 2) * den0
        sum3 = sum3 + iir_y(i + ord + 1) * den1
        sum3 = sum3 + iir_y(i + ord) * den2
        iir_y(i + ord + 3) = -sum3
        y_(yo + i + 3) = sum3
    Next
    For i = i To N - 1
        sum = x_(xo + i)
        For j = 0 To ord - 1
            sum = sum - rden(j) * iir_y(i + j)          ' MULT16_16(rden[j],y[i+j])
        Next
        iir_y(i + ord) = sum                            ' SROUND16(sum,SIG_SHIFT)
        y_(yo + i) = sum
    Next
    For i = 0 To ord - 1
        mem(memo + i) = y_(yo + N - i - 1)
    Next
    ' #endif
End Sub

' (VB) x: x() from xo (in: [0...n-1] samples x); ac: ac() from 0 (out: [0...lag] ac values); window: window()
' from 0, not read when overlap is 0 (the C's NULL)
Public Function celt_autocorr_(x() As Single, ByVal xo As Long, ac() As Single, window() As Single, ByVal overlap As Long, ByVal lag As Long, ByVal n As Long, ByVal arch As Long) As Long
    Dim d As Single, i As Long, k As Long, fastN As Long, shift As Long
    If Not lpc_done Then lpc_init
    fastN = n - lag
    ' ALLOC(xx, n, opus_val16);
    If UBound(autocorr_xx) < n - 1 Then ReDim autocorr_xx(0 To n - 1)
    ' celt_assert(n>0);
    ' celt_assert(overlap>=0);
    ' (VB) xptr is always xx: with overlap 0 a plain copy of x (the C points xptr at x)
    For i = 0 To n - 1
        autocorr_xx(i) = x(xo + i)
    Next
    If overlap <> 0 Then
        For i = 0 To overlap - 1
            autocorr_xx(i) = x(xo + i) * window(i)                      ' MULT16_16_Q15(x[i],window[i])
            autocorr_xx(n - i - 1) = x(xo + n - i - 1) * window(i)      ' MULT16_16_Q15(x[n-i-1],window[i])
        Next
    End If
    shift = 0
    ' #ifdef FIXED_POINT ... #endif
    celt_pitch_xcorr_c autocorr_xx, 0, autocorr_xx, 0, ac, fastN, lag + 1, arch
    For k = 0 To lag
        d = 0
        For i = k + fastN To n - 1
            d = d + autocorr_xx(i) * autocorr_xx(i - k)                 ' MAC16_16(d, xptr[i], xptr[i-k])
        Next
        ac(k) = ac(k) + d
    Next
    ' #ifdef FIXED_POINT ... #endif

    celt_autocorr_ = shift
End Function
