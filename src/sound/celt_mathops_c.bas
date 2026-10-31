Attribute VB_Name = "celt_mathops_c"
' ref/opus/celt (libopus 1.5.2): mathops.c and mathops.h, the float build (FIXED_POINT and FLOAT_APPROX
' undefined) (cluster K, fork O2). mathops.c's FIXED_POINT functions (frac_div32, celt_rsqrt_norm, celt_sqrt,
' _celt_cos_pi_2, celt_cos_norm, celt_rcp) and mathops.h's fixed-point inlines are not compiled in this build.
' The float build's macros are written in place by their users, as the preprocessor does: celt_sqrt(x) is
' Sqr(x), celt_rsqrt(x) and celt_rsqrt_norm(x) 1 / Sqr(x), celt_cos_norm(x) Cos(.5 * PI * x) (vq.c),
' celt_rcp(x) 1 / x, celt_div(a, b) and frac_div32(a, b) a / b, FRAC_MUL16 (bands.c). celt_exp2 is a function
' here (O3's opus_decoder.c calls it too). Not reached by the decoder, so not ported: celt_log2 (amp2Log2's),
' fast_atan2f (stereo_itheta's), celt_maxabs16 / celt_maxabs32 (FIXED_POINT code only).
Option Explicit

'Compute floor(sqrt(_val)) with exact arithmetic.
'  _val must be greater than 0.
'  This has been tested on all possible 32-bit inputs greater than 0.
' (VB) _val is an opus_uint32 below 2^31 here (compute_theta: 8*fm+1 with fm < 2^15)
Public Function isqrt32(ByVal val As Long) As Long
    Dim b As Long, g As Long, bshift As Long, t As Double
    'Uses the second method from
    '   http://www.azillionmonkeys.com/qed/sqroot.html
    '  The main idea is to search for the largest binary digit b such that
    '   (g+b)*(g+b) <= _val, and add it to the solution g.
    g = 0
    bshift = (ec_ilog(val) - 1) \ 2
    b = celt_pow2(bshift)
    Do
        t = CDbl(g * 2 + b) * celt_pow2(bshift)
        If t <= val Then
            g = g + b
            val = val - t
        End If
        b = b \ 2
        bshift = bshift - 1
    Loop While bshift >= 0
    isqrt32 = g
End Function

' #ifndef FIXED_POINT ... #ifndef FLOAT_APPROX
' #define celt_exp2(x) ((float)exp(0.6931471805599453094*(x)))
Public Function celt_exp2(ByVal x As Single) As Single
    celt_exp2 = Exp(0.693147180559945 * x)
End Function
