Attribute VB_Name = "celt_vq_c"
' ref/opus/celt (libopus 1.5.2): vq.c and vq.h, the float build's decoder side (cluster K, fork O2):
' exp_rotation1, exp_rotation, normalise_residual, extract_collapse_mask, alg_unquant, renormalise_vector.
' Not ported (the encoder's): op_pvq_search_c, alg_quant, stereo_itheta. No OVERRIDE_vq_exp_rotation1 /
' OVERRIDE_renormalise_vector (MIPS only).
' (VB) A celt_norm *X is X() and an index Xo; the float build's macros are the plain float operations.
Option Explicit

' (VB) mathops.h: #define PI 3.141592653f; celt_cos_norm(x) is (float)cos((.5f*PI)*(x))
Private Const HALF_PI As Single = 0.5! * 3.141592653!

' (VB) alg_unquant's int iy[] (N <= 176: the widest band at LM 3)
Private iy() As Long
Private vq_done As Boolean

' #ifndef OVERRIDE_vq_exp_rotation1
Private Sub exp_rotation1(X() As Single, ByVal Xo As Long, ByVal len_ As Long, ByVal stride As Long, ByVal c As Single, ByVal s As Single)
    Dim i As Long, ms As Single, Xptr As Long, x1 As Single, x2 As Single
    Xptr = Xo
    ms = -s                                             ' NEG16(s)
    For i = 0 To len_ - stride - 1
        x1 = X(Xptr)
        x2 = X(Xptr + stride)
        X(Xptr + stride) = c * x2 + s * x1              ' EXTRACT16(PSHR32(MAC16_16(MULT16_16(c, x2),  s, x1), 15))
        X(Xptr) = c * x1 + ms * x2                      ' *Xptr++ = EXTRACT16(PSHR32(MAC16_16(MULT16_16(c, x1), ms, x2), 15))
        Xptr = Xptr + 1
    Next
    Xptr = Xo + len_ - 2 * stride - 1
    For i = len_ - 2 * stride - 1 To 0 Step -1
        x1 = X(Xptr)
        x2 = X(Xptr + stride)
        X(Xptr + stride) = c * x2 + s * x1
        X(Xptr) = c * x1 + ms * x2                      ' *Xptr--
        Xptr = Xptr - 1
    Next
End Sub
' #endif /* OVERRIDE_vq_exp_rotation1 */

Public Sub exp_rotation(X() As Single, ByVal Xo As Long, ByVal len_ As Long, ByVal dir_ As Long, ByVal stride As Long, ByVal K As Long, ByVal spread As Long)
    Dim i As Long, c As Single, s As Single, gain As Single, theta As Single, stride2 As Long, factor As Long, t As Single

    If 2 * K >= len_ Or spread = SPREAD_NONE Then Exit Sub
    ' static const int SPREAD_FACTOR[3]={15,10,5};
    Select Case spread - 1
    Case 0: factor = 15
    Case 1: factor = 10
    Case Else: factor = 5
    End Select

    gain = CSng(len_) / CSng(len_ + factor * K)         ' celt_div((opus_val32)MULT16_16(Q15_ONE,len),(opus_val32)(len+factor*K))
    theta = 0.5! * (gain * gain)                        ' HALF16(MULT16_16_Q15(gain,gain))

    t = HALF_PI * theta
    c = Cos(t)                                          ' celt_cos_norm(EXTEND32(theta))
    t = HALF_PI * (1! - theta)
    s = Cos(t)                                          ' celt_cos_norm(EXTEND32(SUB16(Q15ONE,theta)));  sin(theta)

    If len_ >= 8 * stride Then
        stride2 = 1
        ' This is just a simple (equivalent) way of computing sqrt(len/stride) with rounding.
        '   It's basically incrementing long as (stride2+0.5)^2 < len/stride.
        Do While (stride2 * stride2 + stride2) * stride + (stride \ 4) < len_
            stride2 = stride2 + 1
        Loop
    End If
    'NOTE: As a minor optimization, we could be passing around log2(B), not B, for both this and for
    '   extract_collapse_mask().
    len_ = len_ \ stride                                ' celt_udiv(len, stride)
    For i = 0 To stride - 1
        If dir_ < 0 Then
            If stride2 <> 0 Then exp_rotation1 X, Xo + i * len_, len_, stride2, s, c
            exp_rotation1 X, Xo + i * len_, len_, 1, c, s
        Else
            exp_rotation1 X, Xo + i * len_, len_, 1, c, -s
            If stride2 <> 0 Then exp_rotation1 X, Xo + i * len_, len_, stride2, s, -c
        End If
    Next
End Sub

'* Takes the pitch vector and the decoded residual vector, computes the gain
'    that will give ||p+g*y||=1 and mixes the residual with the pitch.
Private Sub normalise_residual(iy_() As Long, X() As Single, ByVal Xo As Long, ByVal N As Long, ByVal Ryy As Single, ByVal gain As Single)
    Dim i As Long, t As Single, g As Single
    ' #ifdef FIXED_POINT ... #endif
    t = Ryy                                             ' VSHR32(Ryy, 2*(k-7))
    g = (1! / Sqr(t)) * gain                            ' MULT16_16_P15(celt_rsqrt_norm(t),gain)

    i = 0
    Do
        X(Xo + i) = g * iy_(i)                          ' EXTRACT16(PSHR32(MULT16_16(g, iy[i]), k+1))
        i = i + 1
    Loop While i < N
End Sub

Private Function extract_collapse_mask(iy_() As Long, ByVal N As Long, ByVal B As Long) As Long
    Dim collapse_mask As Long, N0 As Long, i As Long, j As Long, tmp As Long, bit As Long
    If B <= 1 Then
        extract_collapse_mask = 1
        Exit Function
    End If
    'NOTE: As a minor optimization, we could be passing around log2(B), not B, for both this and for
    '   exp_rotation().
    N0 = N \ B                                          ' celt_udiv(N, B)
    collapse_mask = 0
    bit = 1
    i = 0
    Do
        tmp = 0
        j = 0
        Do
            tmp = tmp Or iy_(i * N0 + j)
            j = j + 1
        Loop While j < N0
        If tmp <> 0 Then collapse_mask = collapse_mask Or bit     ' collapse_mask |= (tmp!=0)<<i;
        bit = bit * 2
        i = i + 1
    Loop While i < B
    extract_collapse_mask = collapse_mask
End Function

' alg_quant: the encoder's (not ported)

'* Decode pulse vector and combine the result with the pitch vector to produce
'    the final normalised signal in the current band.
Public Function alg_unquant(X() As Single, ByVal Xo As Long, ByVal N As Long, ByVal K As Long, ByVal spread As Long, ByVal B As Long, dec As ec_ctx, ByVal gain As Single) As Long
    Dim Ryy As Single
    If Not vq_done Then
        ReDim iy(0 To 255)
        vq_done = True
    End If
    ' celt_assert2(K>0, "alg_unquant() needs at least one pulse");
    ' celt_assert2(N>1, "alg_unquant() needs at least two dimensions");
    Ryy = decode_pulses(iy, 0, N, K, dec)
    normalise_residual iy, X, Xo, N, Ryy, gain
    exp_rotation X, Xo, N, -1, B, K, spread
    alg_unquant = extract_collapse_mask(iy, N, B)
End Function

' #ifndef OVERRIDE_renormalise_vector
Public Sub renormalise_vector(X() As Single, ByVal Xo As Long, ByVal N As Long, ByVal gain As Single, ByVal arch As Long)
    Dim i As Long, E As Single, g As Single, t As Single, xptr As Long
    E = 1E-15! + celt_inner_prod_c(X, Xo, X, Xo, N)     ' EPSILON + celt_inner_prod(X, X, N, arch)
    ' #ifdef FIXED_POINT ... #endif
    t = E                                               ' VSHR32(E, 2*(k-7))
    g = (1! / Sqr(t)) * gain                            ' MULT16_16_P15(celt_rsqrt_norm(t),gain)

    xptr = Xo
    For i = 0 To N - 1
        X(xptr) = g * X(xptr)                           ' EXTRACT16(PSHR32(MULT16_16(g, *xptr), k+1))
        xptr = xptr + 1
    Next
    'return celt_sqrt(E);
End Sub
' #endif /* OVERRIDE_renormalise_vector */

' stereo_itheta: the encoder's (not ported)
