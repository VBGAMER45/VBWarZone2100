Attribute VB_Name = "celt_bands_c"
' ref/opus/celt (libopus 1.5.2): bands.c and bands.h, the float build's decoder side, DISABLE_UPDATE_DRAFT
' undefined, no RESYNTH (cluster K, fork O2). In the functions the encoder shares, its branches are comments
' ("encoder only"); not ported (the encoder's only): hysteresis_decision, compute_band_energies,
' normalise_bands, compute_channel_weights, intensity_stereo, stereo_split, spreading_decision, and the
' theta_rdo part of quant_all_bands (encode && complexity >= 8). MEASURE_NORM_MSE's measure_norm_mse is not
' compiled.
' (VB) Every celt_norm * of the band decoding (X_, Y_, X, Y, lowband, lowband_out, lowband_scratch, norm,
' norm2, the decoder's X) is an index into one array, celt_norm_buf, -1 = NULL: the decoder's X is
' celt_norm_buf from 0, and quant_all_bands' _norm is celt_norm_buf from NORM_BASE (the C ALLOCs it). The
' struct band_ctx has no ec (ec_ctx * is passed as ec, ByRef) and no bandE (the encoder's: NULL here).
' opus_uint32 seed is a Long's bits; celt_lcg_rand's product is exact in a Double.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Public Const SPREAD_NONE As Long = 0
Public Const SPREAD_LIGHT As Long = 1
Public Const SPREAD_NORMAL As Long = 2
Public Const SPREAD_AGGRESSIVE As Long = 3

' (VB) the celt_norm memory: the decoder's X (C*N <= 1920 from 0) and quant_all_bands' _norm (from NORM_BASE:
' 2*(8*78) at most), the C's stack
Public Const NORM_BASE As Long = 1920
Public Const CELT_NORM_BUF_SIZE As Long = 3200
Public celt_norm_buf() As Single

Private Type band_ctx
    encode As Long
    resynth As Long
    m As Long                           ' const CELTMode *m: the mode's handle
    i As Long
    intensity As Long
    spread As Long
    tf_change As Long
    ' ec_ctx *ec: (VB) passed as ec
    remaining_bits As Long              ' opus_int32
    ' const celt_ener *bandE: (VB) the encoder's (NULL in the decoder)
    seed As Long                        ' opus_uint32
    arch As Long
    theta_round As Long
    disable_inv As Long
    avoid_split_noise As Long
End Type

Private Type split_ctx
    inv As Long
    imid As Long
    iside As Long
    delta As Long
    itheta As Long
    qalloc As Long
End Type

' (VB) the static tables: ordery_table, bit_interleave_table, bit_deinterleave_table, compute_qn's exp2_table8;
' deinterleave_hadamard / interleave_hadamard's tmp
Private ordery_table(0 To 29) As Long
Private bit_interleave_table(0 To 15) As Long
Private bit_deinterleave_table(0 To 15) As Long
Private exp2_table8(0 To 7) As Long
Private hadamard_tmp() As Single
Private bands_done As Boolean

' (VB) the tables and celt_norm_buf (called by the decoder's init)
Public Sub celt_bands_init()
    Dim v As Variant, i As Long
    If bands_done Then Exit Sub
    ' Indexing table for converting from natural Hadamard to ordery Hadamard
    '   This is essentially a bit-reversed Gray, on top of which we've added
    '   an inversion of the order because we want the DC at the end rather than
    '   the beginning. The lines are for N=2, 4, 8, 16
    v = Array(1, 0, _
              3, 0, 2, 1, _
              7, 0, 4, 3, 6, 1, 5, 2, _
              15, 0, 8, 7, 12, 3, 11, 4, 14, 1, 9, 6, 13, 2, 10, 5)
    For i = 0 To 29
        ordery_table(i) = v(i)
    Next
    v = Array(0, 1, 1, 1, 2, 3, 3, 3, 2, 3, 3, 3, 2, 3, 3, 3)
    For i = 0 To 15
        bit_interleave_table(i) = v(i)
    Next
    v = Array(&H0, &H3, &HC, &HF, &H30, &H33, &H3C, &H3F, &HC0, &HC3, &HCC, &HCF, &HF0, &HF3, &HFC, &HFF)
    For i = 0 To 15
        bit_deinterleave_table(i) = v(i)
    Next
    v = Array(16384, 17866, 19483, 21247, 23170, 25267, 27554, 30048)
    For i = 0 To 7
        exp2_table8(i) = v(i)
    Next
    ReDim hadamard_tmp(0 To 255)
    ReDim celt_norm_buf(0 To CELT_NORM_BUF_SIZE - 1)
    bands_done = True
End Sub

' (VB) FRAC_MUL16 (mathops.h): ((16384+((opus_int32)(opus_int16)(a)*(opus_int16)(b)))>>15), the >>15 flooring
Private Function FRAC_MUL16(ByVal a As Long, ByVal b As Long) As Long
    Dim t As Long
    a = ((a + 32768) And &HFFFF&) - 32768               ' (opus_int16)
    b = ((b + 32768) And &HFFFF&) - 32768
    t = 16384 + a * b
    FRAC_MUL16 = (t And &HFFFF8000) \ 32768
End Function

' (VB) x >> n of a signed int (floor), 0 <= n <= 30
Private Function SHR_(ByVal x As Long, ByVal n As Long) As Long
    SHR_ = Int(x / celt_pow2(n))
End Function

' hysteresis_decision: the encoder's (not ported)

Public Function celt_lcg_rand(ByVal seed As Long) As Long
    Dim d As Double
    ' return 1664525 * seed + 1013904223;  (VB) modulo 2^32, exact in a Double
    d = seed
    If d < 0 Then d = d + 4294967296#
    d = 1664525# * d + 1013904223#
    d = d - Int(d / 4294967296#) * 4294967296#
    If d >= 2147483648# Then d = d - 4294967296#
    celt_lcg_rand = d
End Function

' This is a cos() approximation designed to be bit-exact on any platform. Bit exactness
'   with this approximation is important because it has an impact on the bit allocation
Public Function bitexact_cos(ByVal x As Long) As Long
    Dim tmp As Long, x2 As Long
    tmp = (4096 + x * x) \ 8192                         ' >>13 (not negative)
    ' celt_sig_assert(tmp<=32767);
    x2 = tmp
    x2 = (32767 - x2) + FRAC_MUL16(x2, (-7651 + FRAC_MUL16(x2, (8277 + FRAC_MUL16(-626, x2)))))
    ' celt_sig_assert(x2<=32766);
    bitexact_cos = 1 + x2
End Function

Public Function bitexact_log2tan(ByVal isin_ As Long, ByVal icos_ As Long) As Long
    Dim lc As Long, ls As Long
    lc = ec_ilog(icos_)
    ls = ec_ilog(isin_)
    icos_ = icos_ * celt_pow2(15 - lc)
    isin_ = isin_ * celt_pow2(15 - ls)
    bitexact_log2tan = (ls - lc) * 2048 _
          + FRAC_MUL16(isin_, FRAC_MUL16(isin_, -2597) + 7932) _
          - FRAC_MUL16(icos_, FRAC_MUL16(icos_, -2597) + 7932)
End Function

' compute_band_energies, normalise_bands: the encoder's (not ported)

' De-normalise the energy to produce the synthesis from the unit-energy bands
' (VB) X: celt_norm_buf from X; freq: freq() from freqo; bandLogE: bandLogE() from bandLogEo
Public Sub denormalise_bands(ByVal m As Long, ByVal X As Long, freq() As Single, ByVal freqo As Long, bandLogE() As Single, ByVal bandLogEo As Long, _
        ByVal start As Long, ByVal end_ As Long, ByVal M_ As Long, ByVal downsample As Long, ByVal silence As Long)
    Dim i As Long, N As Long, bound As Long, f As Long, x_ As Long, j As Long, band_end As Long
    Dim g As Single, lg As Single
    With celt_modes(m)
        N = M_ * .shortMdctSize
        bound = M_ * .eBands(end_)
        If downsample <> 1 Then
            If N \ downsample < bound Then bound = N \ downsample      ' IMIN(bound, N/downsample)
        End If
        If silence <> 0 Then
            bound = 0
            start = 0: end_ = 0
        End If
        f = freqo
        x_ = X + M_ * .eBands(start)
        For i = 0 To M_ * .eBands(start) - 1
            freq(f) = 0
            f = f + 1
        Next
        For i = start To end_ - 1
            j = M_ * .eBands(i)
            band_end = M_ * .eBands(i + 1)
            lg = bandLogE(bandLogEo + i) + eMeans(i)    ' SATURATE16(ADD32(bandLogE[i], SHL32((opus_val32)eMeans[i],6)))
            ' #ifndef FIXED_POINT
            If lg > 32! Then lg = 32!                   ' MIN32(32.f, lg)
            g = celt_exp2(lg)
            ' #else ... #endif
            ' Be careful of the fixed-point "else" just above when changing this code
            For j = j To band_end - 1
                freq(f) = celt_norm_buf(x_) * g         ' *f++ = SHR32(MULT16_16(*x++, g), shift);
                f = f + 1
                x_ = x_ + 1
            Next
        Next
        ' celt_assert(start <= end);
        For i = bound To N - 1                          ' OPUS_CLEAR(&freq[bound], N-bound);
            freq(freqo + i) = 0
        Next
    End With
End Sub

' This prevents energy collapse for transients with multiple short MDCTs
' (VB) X_: celt_norm_buf from X_; logE, prev1logE, prev2logE: arrays from 0; pulses: pulses() from 0
Public Sub anti_collapse(ByVal m As Long, ByVal X_ As Long, collapse_masks() As Byte, ByVal LM As Long, ByVal C As Long, ByVal size As Long, _
        ByVal start As Long, ByVal end_ As Long, logE() As Single, prev1logE() As Single, _
        prev2logE() As Single, pulses() As Long, ByVal seed As Long, ByVal arch As Long)
    Dim c_ As Long, i As Long, j As Long, k As Long, N0 As Long, thresh As Single, sqrt_1 As Single, depth As Long
    Dim X As Long, prev1 As Single, prev2 As Single, Ediff As Single, r As Single, renormalize As Long
    Dim nbEBands As Long, M_ As Long, bit As Long
    nbEBands = celt_modes(m).nbEBands
    M_ = celt_pow2(LM)
    For i = start To end_ - 1
        N0 = celt_modes(m).eBands(i + 1) - celt_modes(m).eBands(i)
        ' depth in 1/8 bits
        ' celt_sig_assert(pulses[i]>=0);
        depth = ((1 + pulses(i)) \ N0) \ M_             ' celt_udiv(1+pulses[i], (m->eBands[i+1]-m->eBands[i]))>>LM

        ' #ifdef FIXED_POINT ... #else
        thresh = 0.5! * celt_exp2(-0.125! * depth)
        sqrt_1 = 1! / Sqr(N0 * M_)                      ' celt_rsqrt(N0<<LM)
        ' #endif

        c_ = 0
        Do
            renormalize = 0
            prev1 = prev1logE(c_ * nbEBands + i)
            prev2 = prev2logE(c_ * nbEBands + i)
            If C = 1 Then
                If prev1logE(nbEBands + i) > prev1 Then prev1 = prev1logE(nbEBands + i)  ' MAX16(prev1,prev1logE[m->nbEBands+i])
                If prev2logE(nbEBands + i) > prev2 Then prev2 = prev2logE(nbEBands + i)
            End If
            If prev1 < prev2 Then
                Ediff = logE(c_ * nbEBands + i) - prev1     ' EXTEND32(logE[c*m->nbEBands+i])-EXTEND32(MIN16(prev1,prev2))
            Else
                Ediff = logE(c_ * nbEBands + i) - prev2
            End If
            If Ediff < 0 Then Ediff = 0                 ' MAX32(0, Ediff)

            ' #ifdef FIXED_POINT ... #else
            ' r needs to be multiplied by 2 or 2*sqrt(2) depending on LM because
            '   short blocks don't have the same energy as long
            r = 2! * celt_exp2(-Ediff)
            If LM = 3 Then r = r * 1.41421356!
            If thresh < r Then r = thresh               ' MIN16(thresh, r)
            r = r * sqrt_1
            ' #endif
            X = X_ + c_ * size + celt_modes(m).eBands(i) * M_
            bit = 1
            For k = 0 To M_ - 1
                ' Detect collapse
                If (collapse_masks(i * C + c_) And bit) = 0 Then
                    ' Fill with noise
                    For j = 0 To N0 - 1
                        seed = celt_lcg_rand(seed)
                        If (seed And &H8000&) <> 0 Then
                            celt_norm_buf(X + j * M_ + k) = r
                        Else
                            celt_norm_buf(X + j * M_ + k) = -r
                        End If
                    Next
                    renormalize = 1
                End If
                bit = bit * 2
            Next
            ' We just added some energy, so we need to renormalise
            If renormalize <> 0 Then renormalise_vector celt_norm_buf, X, N0 * M_, 1!, arch
            c_ = c_ + 1
        Loop While c_ < C
    Next
End Sub

' compute_channel_weights, intensity_stereo, stereo_split: the encoder's (not ported)

' (VB) X, Y: celt_norm_buf from X, Y
Private Sub stereo_merge(ByVal X As Long, ByVal Y As Long, ByVal mid As Single, ByVal N As Long, ByVal arch As Long)
    Dim j As Long, xp As Single, side As Single, El As Single, Er As Single, mid2 As Single
    Dim t As Single, lgain As Single, rgain As Single, l As Single, r As Single
    xp = 0: side = 0
    ' Compute the norm of X+Y and X-Y as |X|^2 + |Y|^2 +/- sum(xy)
    dual_inner_prod_c celt_norm_buf, Y, celt_norm_buf, X, celt_norm_buf, Y, N, xp, side
    ' Compensating for the mid normalization
    xp = mid * xp                                       ' MULT16_32_Q15(mid, xp)
    ' mid and side are in Q15, not Q14 like X and Y
    mid2 = mid                                          ' SHR16(mid, 1)
    El = mid2 * mid2 + side - 2 * xp
    Er = mid2 * mid2 + side + 2 * xp
    If Er < 0.0006! Or El < 0.0006! Then                ' QCONST32(6e-4f, 28)
        CopyMemory celt_norm_buf(Y), celt_norm_buf(X), N * 4    ' OPUS_COPY(Y, X, N)
        Exit Sub
    End If

    ' #ifdef FIXED_POINT ... #endif
    t = El                                              ' VSHR32(El, (kl-7)<<1)
    lgain = 1! / Sqr(t)                                 ' celt_rsqrt_norm(t)
    t = Er
    rgain = 1! / Sqr(t)

    ' #ifdef FIXED_POINT ... #endif

    For j = 0 To N - 1
        ' Apply mid scaling (side is already scaled)
        l = mid * celt_norm_buf(X + j)                  ' MULT16_16_P15(mid, X[j])
        r = celt_norm_buf(Y + j)
        celt_norm_buf(X + j) = lgain * (l - r)          ' EXTRACT16(PSHR32(MULT16_16(lgain, SUB16(l,r)), kl+1))
        celt_norm_buf(Y + j) = rgain * (l + r)
    Next
End Sub

' spreading_decision: the encoder's (not ported)

' (VB) X: celt_norm_buf from X
Private Sub deinterleave_hadamard(ByVal X As Long, ByVal N0 As Long, ByVal stride As Long, ByVal hadamard As Long)
    Dim i As Long, j As Long, N As Long, ordery As Long
    N = N0 * stride
    ' ALLOC(tmp, N, celt_norm);
    ' celt_assert(stride>0);
    If hadamard <> 0 Then
        ordery = stride - 2                             ' const int *ordery = ordery_table+stride-2;
        For i = 0 To stride - 1
            For j = 0 To N0 - 1
                hadamard_tmp(ordery_table(ordery + i) * N0 + j) = celt_norm_buf(X + j * stride + i)
            Next
        Next
    Else
        For i = 0 To stride - 1
            For j = 0 To N0 - 1
                hadamard_tmp(i * N0 + j) = celt_norm_buf(X + j * stride + i)
            Next
        Next
    End If
    CopyMemory celt_norm_buf(X), hadamard_tmp(0), N * 4     ' OPUS_COPY(X, tmp, N);
End Sub

' (VB) X: celt_norm_buf from X
Private Sub interleave_hadamard(ByVal X As Long, ByVal N0 As Long, ByVal stride As Long, ByVal hadamard As Long)
    Dim i As Long, j As Long, N As Long, ordery As Long
    N = N0 * stride
    ' ALLOC(tmp, N, celt_norm);
    If hadamard <> 0 Then
        ordery = stride - 2                             ' const int *ordery = ordery_table+stride-2;
        For i = 0 To stride - 1
            For j = 0 To N0 - 1
                hadamard_tmp(j * stride + i) = celt_norm_buf(X + ordery_table(ordery + i) * N0 + j)
            Next
        Next
    Else
        For i = 0 To stride - 1
            For j = 0 To N0 - 1
                hadamard_tmp(j * stride + i) = celt_norm_buf(X + i * N0 + j)
            Next
        Next
    End If
    CopyMemory celt_norm_buf(X), hadamard_tmp(0), N * 4     ' OPUS_COPY(X, tmp, N);
End Sub

' (VB) X: celt_norm_buf from X
Public Sub haar1(ByVal X As Long, ByVal N0 As Long, ByVal stride As Long)
    Dim i As Long, j As Long, tmp1 As Single, tmp2 As Single, k1 As Long, k2 As Long
    N0 = N0 \ 2
    For i = 0 To stride - 1
        For j = 0 To N0 - 1
            k1 = X + stride * 2 * j + i
            k2 = X + stride * (2 * j + 1) + i
            tmp1 = 0.70710678! * celt_norm_buf(k1)      ' MULT16_16(QCONST16(.70710678f,15), X[stride*2*j+i])
            tmp2 = 0.70710678! * celt_norm_buf(k2)
            celt_norm_buf(k1) = tmp1 + tmp2             ' EXTRACT16(PSHR32(ADD32(tmp1, tmp2), 15))
            celt_norm_buf(k2) = tmp1 - tmp2
        Next
    Next
End Sub

Private Function compute_qn(ByVal N As Long, ByVal b As Long, ByVal offset As Long, ByVal pulse_cap As Long, ByVal stereo As Long) As Long
    Dim qn As Long, qb As Long, N2 As Long
    N2 = 2 * N - 1
    If stereo <> 0 And N = 2 Then N2 = N2 - 1
    ' The upper limit ensures that in a stereo split with itheta==16384, we'll
    '    always have enough bits left over to code at least one pulse in the
    '    side; otherwise it would collapse, since it doesn't get folded.
    qb = (b + N2 * offset) \ N2                         ' celt_sudiv(b+N2*offset, N2)
    If b - pulse_cap - (4 * 8) < qb Then qb = b - pulse_cap - (4 * 8)   ' IMIN(b-pulse_cap-(4<<BITRES), qb)

    If 8 * 8 < qb Then qb = 8 * 8                       ' IMIN(8<<BITRES, qb)

    If qb < 4 Then                                      ' (1<<BITRES>>1)
        qn = 1
    Else
        qn = exp2_table8(qb And 7) \ celt_pow2(14 - (qb \ 8))
        qn = ((qn + 1) \ 2) * 2                         ' (qn+1)>>1<<1
    End If
    ' celt_assert(qn <= 256);
    compute_qn = qn
End Function

' (VB) X, Y: celt_norm_buf offsets (the encoder's use only); b and fill ByRef (int *b, int *fill)
Private Sub compute_theta(ctx As band_ctx, sctx As split_ctx, ByVal X As Long, ByVal Y As Long, ByVal N As Long, b As Long, ByVal B_ As Long, ByVal B0 As Long, _
        ByVal LM As Long, ByVal stereo As Long, fill As Long, ec As ec_ctx)
    Dim qn As Long, itheta As Long, delta As Long, imid As Long, iside As Long, qalloc As Long, pulse_cap As Long, offset As Long
    Dim tell As Long, inv As Long, encode As Long, m As Long, i As Long, intensity As Long
    Dim p0 As Long, x_ As Long, x0 As Long, ft As Long, fs As Long, fl As Long, fm As Long, fh As Long

    encode = ctx.encode
    m = ctx.m
    i = ctx.i
    intensity = ctx.intensity
    ' ec = ctx->ec; bandE = ctx->bandE;

    ' Decide on the resolution to give to the split parameter theta
    pulse_cap = celt_modes(m).logN(i) + LM * 8          ' m->logN[i]+LM*(1<<BITRES)
    If stereo <> 0 And N = 2 Then
        offset = SHR_(pulse_cap, 1) - QTHETA_OFFSET_TWOPHASE
    Else
        offset = SHR_(pulse_cap, 1) - QTHETA_OFFSET
    End If
    qn = compute_qn(N, b, offset, pulse_cap, stereo)
    If stereo <> 0 And i >= intensity Then qn = 1
    If encode <> 0 Then
        ' theta is the atan() of the ratio between the (normalized)
        '   side and mid. ...  (encoder only: itheta = stereo_itheta(X, Y, stereo, N, ctx->arch))
    End If
    tell = ec_tell_frac(ec)
    If qn <> 1 Then
        If encode <> 0 Then
            ' (encoder only: the quantisation of itheta, theta_round, avoid_split_noise)
        End If
        ' Entropy coding of the angle. We use a uniform pdf for the
        '   time split, a step for stereo, and a triangular one for the rest.
        If stereo <> 0 And N > 2 Then
            p0 = 3
            x_ = itheta
            x0 = qn \ 2
            ft = p0 * (x0 + 1) + x0
            ' Use a probability of p0 up to itheta=8192 and then use 1 after
            If encode <> 0 Then
                ' (encoder only: ec_encode)
            Else
                fs = ec_decode(ec, ft)
                If fs < (x0 + 1) * p0 Then
                    x_ = fs \ p0
                Else
                    x_ = x0 + 1 + (fs - (x0 + 1) * p0)
                End If
                If x_ <= x0 Then
                    fl = p0 * x_
                    fh = p0 * (x_ + 1)
                Else
                    fl = (x_ - 1 - x0) + (x0 + 1) * p0
                    fh = (x_ - x0) + (x0 + 1) * p0
                End If
                ec_dec_update ec, fl, fh, ft
                itheta = x_
            End If
        ElseIf B0 > 1 Or stereo <> 0 Then
            ' Uniform pdf
            If encode <> 0 Then
                ' (encoder only: ec_enc_uint)
            Else
                itheta = ec_dec_uint(ec, qn + 1)
            End If
        Else
            fs = 1
            ft = ((qn \ 2) + 1) * ((qn \ 2) + 1)
            If encode <> 0 Then
                ' (encoder only: ec_encode)
            Else
                ' Triangular pdf
                fl = 0
                fm = ec_decode(ec, ft)

                If fm < (((qn \ 2) * ((qn \ 2) + 1)) \ 2) Then
                    itheta = (isqrt32(8 * fm + 1) - 1) \ 2
                    fs = itheta + 1
                    fl = (itheta * (itheta + 1)) \ 2
                Else
                    itheta = (2 * (qn + 1) - isqrt32(8 * (ft - fm - 1) + 1)) \ 2
                    fs = qn + 1 - itheta
                    fl = ft - (((qn + 1 - itheta) * (qn + 2 - itheta)) \ 2)
                End If

                ec_dec_update ec, fl, fl + fs, ft
            End If
        End If
        ' celt_assert(itheta>=0);
        itheta = (itheta * 16384) \ qn                  ' celt_udiv((opus_int32)itheta*16384, qn)
        If encode <> 0 And stereo <> 0 Then
            ' (encoder only: intensity_stereo / stereo_split)
        End If
        ' NOTE: Renormalising X and Y *may* help fixed-point a bit at very high rate.
        '         Let's do that at higher complexity
    ElseIf stereo <> 0 Then
        If encode <> 0 Then
            ' (encoder only: inv, intensity_stereo)
        End If
        If b > 2 * 8 And ctx.remaining_bits > 2 * 8 Then    ' *b>2<<BITRES && ctx->remaining_bits > 2<<BITRES
            If encode <> 0 Then
                ' (encoder only: ec_enc_bit_logp)
            Else
                inv = ec_dec_bit_logp(ec, 2)
            End If
        Else
            inv = 0
        End If
        ' inv flag override to avoid problems with downmixing.
        If ctx.disable_inv <> 0 Then inv = 0
        itheta = 0
    End If
    qalloc = ec_tell_frac(ec) - tell
    b = b - qalloc

    If itheta = 0 Then
        imid = 32767
        iside = 0
        fill = fill And (celt_pow2(B_) - 1)                  ' *fill &= (1<<B)-1;
        delta = -16384
    ElseIf itheta = 16384 Then
        imid = 0
        iside = 32767
        fill = fill And ((celt_pow2(B_) - 1) * celt_pow2(B_))     ' *fill &= ((1<<B)-1)<<B;
        delta = 16384
    Else
        imid = bitexact_cos(itheta)
        iside = bitexact_cos(16384 - itheta)
        ' This is the mid vs side allocation that minimizes squared error
        '   in that band.
        delta = FRAC_MUL16((N - 1) * 128, bitexact_log2tan(iside, imid))
    End If

    sctx.inv = inv
    sctx.imid = imid
    sctx.iside = iside
    sctx.delta = delta
    sctx.itheta = itheta
    sctx.qalloc = qalloc
End Sub

' (VB) X, Y (-1 = NULL), lowband_out (-1 = NULL): celt_norm_buf offsets
Private Function quant_band_n1(ctx As band_ctx, ByVal X As Long, ByVal Y As Long, ByVal lowband_out As Long, ec As ec_ctx) As Long
    Dim c_ As Long, stereo As Long, x_ As Long, encode As Long, sign As Long

    encode = ctx.encode
    ' ec = ctx->ec;

    If Y <> -1 Then stereo = 1
    x_ = X
    c_ = 0
    Do
        sign = 0
        If ctx.remaining_bits >= 8 Then                 ' 1<<BITRES
            If encode <> 0 Then
                ' (encoder only: sign = x[0]<0; ec_enc_bits(ec, sign, 1);)
            Else
                sign = ec_dec_bits(ec, 1)
            End If
            ctx.remaining_bits = ctx.remaining_bits - 8
        End If
        If ctx.resynth <> 0 Then
            If sign <> 0 Then celt_norm_buf(x_) = -1! Else celt_norm_buf(x_) = 1!     ' sign ? -NORM_SCALING : NORM_SCALING
        End If
        x_ = Y
        c_ = c_ + 1
    Loop While c_ < 1 + stereo
    If lowband_out <> -1 Then celt_norm_buf(lowband_out) = celt_norm_buf(X)    ' SHR16(X[0],4)
    quant_band_n1 = 1
End Function

' This function is responsible for encoding and decoding a mono partition.
'   It can split the band in two and transmit the energy difference with
'   the two half-bands. It can be called recursively so bands can end up being
'   split in 8 parts.
' (VB) X, lowband (-1 = NULL): celt_norm_buf offsets
Private Function quant_partition(ctx As band_ctx, ByVal X As Long, ByVal N As Long, ByVal b As Long, ByVal B_ As Long, ByVal lowband As Long, _
        ByVal LM As Long, ByVal gain As Single, ByVal fill As Long, ec As ec_ctx) As Long
    Dim cache As Long, q As Long, curr_bits As Long, imid As Long, iside As Long, B0 As Long, mid As Single, side As Single
    Dim cm As Long, Y As Long, encode As Long, m As Long, i As Long, spread As Long, split As Boolean
    Dim mbits As Long, sbits As Long, delta As Long, itheta As Long, qalloc As Long, sctx As split_ctx
    Dim next_lowband2 As Long, rebalance As Long, K As Long, j As Long, cm_mask As Long, tmp As Single
    Dim sd As Double, s As Long

    imid = 0: iside = 0
    B0 = B_
    mid = 0: side = 0
    cm = 0
    Y = -1
    encode = ctx.encode
    m = ctx.m
    i = ctx.i
    spread = ctx.spread
    ' ec = ctx->ec;

    ' If we need 1.5 more bit than we can produce, split the band in two.
    cache = celt_modes(m).cache.index((LM + 1) * celt_modes(m).nbEBands + i)   ' cache = m->cache.bits + m->cache.index[...]
    If LM <> -1 Then                                    ' LM != -1 && b > cache[cache[0]]+12 && N>2
        If b > celt_modes(m).cache.bits(cache + celt_modes(m).cache.bits(cache)) + 12 Then
            If N > 2 Then split = True
        End If
    End If
    If split Then
        next_lowband2 = -1

        N = N \ 2
        Y = X + N
        LM = LM - 1
        If B_ = 1 Then fill = (fill And 1) Or (fill * 2)
        B_ = (B_ + 1) \ 2

        compute_theta ctx, sctx, X, Y, N, b, B_, B0, LM, 0, fill, ec
        imid = sctx.imid
        iside = sctx.iside
        delta = sctx.delta
        itheta = sctx.itheta
        qalloc = sctx.qalloc
        ' #ifdef FIXED_POINT ... #else
        mid = (1! / 32768) * imid
        side = (1! / 32768) * iside
        ' #endif

        ' Give more bits to low-energy MDCTs than they would otherwise deserve
        If B0 > 1 And (itheta And &H3FFF) <> 0 Then
            If itheta > 8192 Then
                ' Rough approximation for pre-echo masking
                delta = delta - SHR_(delta, 4 - LM)
            Else
                ' Corresponds to a forward-masking slope of 1.5 dB per 10 ms
                delta = delta + ((N * 8) \ celt_pow2(5 - LM))  ' delta + (N<<BITRES>>(5-LM))
                If delta > 0 Then delta = 0             ' IMIN(0, ...)
            End If
        End If
        mbits = (b - delta) \ 2
        If b < mbits Then mbits = b
        If mbits < 0 Then mbits = 0                     ' IMAX(0, IMIN(b, (b-delta)/2))
        sbits = b - mbits
        ctx.remaining_bits = ctx.remaining_bits - qalloc

        If lowband <> -1 Then next_lowband2 = lowband + N  ' >32-bit split case

        rebalance = ctx.remaining_bits
        If mbits >= sbits Then
            cm = quant_partition(ctx, X, N, mbits, B_, lowband, LM, gain * mid, fill, ec)  ' MULT16_16_P15(gain,mid)
            rebalance = mbits - (rebalance - ctx.remaining_bits)
            If rebalance > 3 * 8 And itheta <> 0 Then sbits = sbits + rebalance - (3 * 8)
            cm = cm Or (quant_partition(ctx, Y, N, sbits, B_, next_lowband2, LM, gain * side, fill \ celt_pow2(B_), ec) * celt_pow2(B0 \ 2))
        Else
            cm = quant_partition(ctx, Y, N, sbits, B_, next_lowband2, LM, gain * side, fill \ celt_pow2(B_), ec) * celt_pow2(B0 \ 2)
            rebalance = sbits - (rebalance - ctx.remaining_bits)
            If rebalance > 3 * 8 And itheta <> 16384 Then mbits = mbits + rebalance - (3 * 8)
            cm = cm Or quant_partition(ctx, X, N, mbits, B_, lowband, LM, gain * mid, fill, ec)
        End If
    Else
        ' This is the basic no-split case
        q = bits2pulses(m, i, LM, b)
        curr_bits = pulses2bits(m, i, LM, q)
        ctx.remaining_bits = ctx.remaining_bits - curr_bits

        ' Ensures we can never bust the budget
        Do While ctx.remaining_bits < 0 And q > 0
            ctx.remaining_bits = ctx.remaining_bits + curr_bits
            q = q - 1
            curr_bits = pulses2bits(m, i, LM, q)
            ctx.remaining_bits = ctx.remaining_bits - curr_bits
        Loop

        If q <> 0 Then
            K = get_pulses(q)

            ' Finally do the actual quantization
            If encode <> 0 Then
                ' (encoder only: cm = alg_quant(X, N, K, spread, B, ec, gain, ctx->resynth, ctx->arch);)
            Else
                cm = alg_unquant(celt_norm_buf, X, N, K, spread, B_, ec, gain)
            End If
        Else
            ' If there's no pulse, fill the band anyway
            If ctx.resynth <> 0 Then
                ' B can be as large as 16, so this shift might overflow an int on a
                '   16-bit platform; use a long to get defined behavior.
                cm_mask = celt_pow2(B_) - 1                  ' (unsigned)(1UL<<B)-1
                fill = fill And cm_mask
                If fill = 0 Then
                    For j = 0 To N - 1                  ' OPUS_CLEAR(X, N);
                        celt_norm_buf(X + j) = 0
                    Next
                Else
                    ' (VB) ctx->seed = celt_lcg_rand(ctx->seed), written out with the seed in a Double
                    sd = ctx.seed
                    If sd < 0 Then sd = sd + 4294967296#
                    If lowband = -1 Then
                        ' Noise
                        For j = 0 To N - 1
                            sd = 1664525# * sd + 1013904223#
                            sd = sd - Int(sd / 4294967296#) * 4294967296#
                            If sd >= 2147483648# Then s = sd - 4294967296# Else s = sd
                            celt_norm_buf(X + j) = (s And &HFFF00000) \ &H100000    ' (celt_norm)((opus_int32)ctx->seed>>20)
                        Next
                        cm = cm_mask
                    Else
                        ' Folded spectrum
                        For j = 0 To N - 1
                            sd = 1664525# * sd + 1013904223#
                            sd = sd - Int(sd / 4294967296#) * 4294967296#
                            If sd >= 2147483648# Then s = sd - 4294967296# Else s = sd
                            ' About 48 dB below the "normal" folding level
                            tmp = 1! / 256                  ' QCONST16(1.0f/256, 10)
                            If (s And &H8000&) = 0 Then tmp = -tmp
                            celt_norm_buf(X + j) = celt_norm_buf(lowband + j) + tmp
                        Next
                        cm = fill
                    End If
                    If sd >= 2147483648# Then ctx.seed = sd - 4294967296# Else ctx.seed = sd
                    renormalise_vector celt_norm_buf, X, N, gain, ctx.arch
                End If
            End If
        End If
    End If

    quant_partition = cm
End Function

' This function is responsible for encoding and decoding a band for the mono case.
' (VB) X, lowband, lowband_out, lowband_scratch (-1 = NULL): celt_norm_buf offsets
Private Function quant_band(ctx As band_ctx, ByVal X As Long, ByVal N As Long, ByVal b As Long, ByVal B_ As Long, ByVal lowband As Long, _
        ByVal LM As Long, ByVal lowband_out As Long, ByVal gain As Single, ByVal lowband_scratch As Long, ByVal fill As Long, ec As ec_ctx) As Long
    Dim N0 As Long, N_B As Long, N_B0 As Long, B0 As Long, time_divide As Long, recombine As Long, longBlocks As Long
    Dim cm As Long, k As Long, encode As Long, tf_change As Long, j As Long, n_ As Single

    N0 = N
    N_B = N
    B0 = B_
    time_divide = 0
    recombine = 0
    cm = 0
    encode = ctx.encode
    tf_change = ctx.tf_change

    If B0 = 1 Then longBlocks = 1

    N_B = N_B \ B_                                      ' celt_udiv(N_B, B)

    ' Special case for one sample
    If N = 1 Then
        quant_band = quant_band_n1(ctx, X, -1, lowband_out, ec)
        Exit Function
    End If

    If tf_change > 0 Then recombine = tf_change
    ' Band recombining to increase frequency resolution

    If lowband_scratch <> -1 And lowband <> -1 Then
        If recombine <> 0 Or ((N_B And 1) = 0 And tf_change < 0) Or B0 > 1 Then
            CopyMemory celt_norm_buf(lowband_scratch), celt_norm_buf(lowband), N * 4  ' OPUS_COPY(lowband_scratch, lowband, N);
            lowband = lowband_scratch
        End If
    End If

    For k = 0 To recombine - 1
        ' static const unsigned char bit_interleave_table[16]={...}: (VB) bands' table
        If encode <> 0 Then
            ' (encoder only: haar1(X, N>>k, 1<<k);)
        End If
        If lowband <> -1 Then haar1 lowband, N \ celt_pow2(k), celt_pow2(k)
        fill = bit_interleave_table(fill And &HF) Or (bit_interleave_table(fill \ 16) * 4)
    Next
    B_ = B_ \ celt_pow2(recombine)
    N_B = N_B * celt_pow2(recombine)

    ' Increasing the time resolution
    Do While (N_B And 1) = 0 And tf_change < 0
        If encode <> 0 Then
            ' (encoder only: haar1(X, N_B, B);)
        End If
        If lowband <> -1 Then haar1 lowband, N_B, B_
        fill = fill Or (fill * celt_pow2(B_))
        B_ = B_ * 2
        N_B = N_B \ 2
        time_divide = time_divide + 1
        tf_change = tf_change + 1
    Loop
    B0 = B_
    N_B0 = N_B

    ' Reorganize the samples in time order instead of frequency order
    If B0 > 1 Then
        If encode <> 0 Then
            ' (encoder only: deinterleave_hadamard(X, N_B>>recombine, B0<<recombine, longBlocks);)
        End If
        If lowband <> -1 Then deinterleave_hadamard lowband, N_B \ celt_pow2(recombine), B0 * celt_pow2(recombine), longBlocks
    End If

    cm = quant_partition(ctx, X, N, b, B_, lowband, LM, gain, fill, ec)

    ' This code is used by the decoder and by the resynthesis-enabled encoder
    If ctx.resynth <> 0 Then
        ' Undo the sample reorganization going from time order to frequency order
        If B0 > 1 Then interleave_hadamard X, N_B \ celt_pow2(recombine), B0 * celt_pow2(recombine), longBlocks

        ' Undo time-freq changes that we did earlier
        N_B = N_B0
        B_ = B0
        For k = 0 To time_divide - 1
            B_ = B_ \ 2
            N_B = N_B * 2
            cm = cm Or (cm \ celt_pow2(B_))
            haar1 X, N_B, B_
        Next

        For k = 0 To recombine - 1
            ' static const unsigned char bit_deinterleave_table[16]={...}: (VB) bands' table
            cm = bit_deinterleave_table(cm)
            haar1 X, N0 \ celt_pow2(k), celt_pow2(k)
        Next
        B_ = B_ * celt_pow2(recombine)

        ' Scale output for later folding
        If lowband_out <> -1 Then
            n_ = Sqr(N0)                                ' celt_sqrt(SHL32(EXTEND32(N0),22))
            For j = 0 To N0 - 1
                celt_norm_buf(lowband_out + j) = n_ * celt_norm_buf(X + j)     ' MULT16_16_Q15(n,X[j])
            Next
        End If
        cm = cm And (celt_pow2(B_) - 1)
    End If
    quant_band = cm
End Function

' This function is responsible for encoding and decoding a band for the stereo case.
' (VB) X, Y, lowband, lowband_out, lowband_scratch (-1 = NULL): celt_norm_buf offsets
Private Function quant_band_stereo(ctx As band_ctx, ByVal X As Long, ByVal Y As Long, ByVal N As Long, ByVal b As Long, ByVal B_ As Long, ByVal lowband As Long, _
        ByVal LM As Long, ByVal lowband_out As Long, ByVal lowband_scratch As Long, ByVal fill As Long, ec As ec_ctx) As Long
    Dim imid As Long, iside As Long, inv As Long, mid As Single, side As Single, cm As Long, mbits As Long, sbits As Long
    Dim delta As Long, itheta As Long, qalloc As Long, sctx As split_ctx, orig_fill As Long, encode As Long
    Dim c_ As Long, sign As Long, x2 As Long, y2 As Long, tmp As Single, rebalance As Long, j As Long

    encode = ctx.encode
    ' ec = ctx->ec;

    ' Special case for one sample
    If N = 1 Then
        quant_band_stereo = quant_band_n1(ctx, X, Y, lowband_out, ec)
        Exit Function
    End If

    orig_fill = fill

    compute_theta ctx, sctx, X, Y, N, b, B_, B_, LM, 1, fill, ec
    inv = sctx.inv
    imid = sctx.imid
    iside = sctx.iside
    delta = sctx.delta
    itheta = sctx.itheta
    qalloc = sctx.qalloc
    ' #ifdef FIXED_POINT ... #else
    mid = (1! / 32768) * imid
    side = (1! / 32768) * iside
    ' #endif

    ' This is a special case for N=2 that only works for stereo and takes
    '   advantage of the fact that mid and side are orthogonal to encode
    '   the side with just one bit.
    If N = 2 Then
        mbits = b
        sbits = 0
        ' Only need one bit for the side.
        If itheta <> 0 And itheta <> 16384 Then sbits = 8   ' 1<<BITRES
        mbits = mbits - sbits
        If itheta > 8192 Then c_ = 1 Else c_ = 0
        ctx.remaining_bits = ctx.remaining_bits - (qalloc + sbits)

        If c_ <> 0 Then
            x2 = Y: y2 = X
        Else
            x2 = X: y2 = Y
        End If
        If sbits <> 0 Then
            If encode <> 0 Then
                ' (encoder only: Here we only need to encode a sign for the side.)
            Else
                sign = ec_dec_bits(ec, 1)
            End If
        End If
        sign = 1 - 2 * sign
        ' We use orig_fill here because we want to fold the side, but if
        '   itheta==16384, we'll have cleared the low bits of fill.
        cm = quant_band(ctx, x2, N, mbits, B_, lowband, LM, lowband_out, 1!, lowband_scratch, orig_fill, ec)
        ' We don't split N=2 bands, so cm is either 1 or 0 (for a fold-collapse),
        '   and there's no need to worry about mixing with the other channel.
        celt_norm_buf(y2) = -sign * celt_norm_buf(x2 + 1)
        celt_norm_buf(y2 + 1) = sign * celt_norm_buf(x2)
        If ctx.resynth <> 0 Then
            celt_norm_buf(X) = mid * celt_norm_buf(X)               ' MULT16_16_Q15(mid, X[0])
            celt_norm_buf(X + 1) = mid * celt_norm_buf(X + 1)
            celt_norm_buf(Y) = side * celt_norm_buf(Y)
            celt_norm_buf(Y + 1) = side * celt_norm_buf(Y + 1)
            tmp = celt_norm_buf(X)
            celt_norm_buf(X) = tmp - celt_norm_buf(Y)
            celt_norm_buf(Y) = tmp + celt_norm_buf(Y)
            tmp = celt_norm_buf(X + 1)
            celt_norm_buf(X + 1) = tmp - celt_norm_buf(Y + 1)
            celt_norm_buf(Y + 1) = tmp + celt_norm_buf(Y + 1)
        End If
    Else
        ' "Normal" split code
        mbits = (b - delta) \ 2
        If b < mbits Then mbits = b
        If mbits < 0 Then mbits = 0                     ' IMAX(0, IMIN(b, (b-delta)/2))
        sbits = b - mbits
        ctx.remaining_bits = ctx.remaining_bits - qalloc

        rebalance = ctx.remaining_bits
        If mbits >= sbits Then
            ' In stereo mode, we do not apply a scaling to the mid because we need the normalized
            '   mid for folding later.
            cm = quant_band(ctx, X, N, mbits, B_, lowband, LM, lowband_out, 1!, lowband_scratch, fill, ec)
            rebalance = mbits - (rebalance - ctx.remaining_bits)
            If rebalance > 3 * 8 And itheta <> 0 Then sbits = sbits + rebalance - (3 * 8)

            ' For a stereo split, the high bits of fill are always zero, so no
            '   folding will be done to the side.
            cm = cm Or quant_band(ctx, Y, N, sbits, B_, -1, LM, -1, side, -1, fill \ celt_pow2(B_), ec)
        Else
            ' For a stereo split, the high bits of fill are always zero, so no
            '   folding will be done to the side.
            cm = quant_band(ctx, Y, N, sbits, B_, -1, LM, -1, side, -1, fill \ celt_pow2(B_), ec)
            rebalance = sbits - (rebalance - ctx.remaining_bits)
            If rebalance > 3 * 8 And itheta <> 16384 Then mbits = mbits + rebalance - (3 * 8)
            ' In stereo mode, we do not apply a scaling to the mid because we need the normalized
            '   mid for folding later.
            cm = cm Or quant_band(ctx, X, N, mbits, B_, lowband, LM, lowband_out, 1!, lowband_scratch, fill, ec)
        End If
    End If

    ' This code is used by the decoder and by the resynthesis-enabled encoder
    If ctx.resynth <> 0 Then
        If N <> 2 Then stereo_merge X, Y, mid, N, ctx.arch
        If inv <> 0 Then
            For j = 0 To N - 1
                celt_norm_buf(Y + j) = -celt_norm_buf(Y + j)
            Next
        End If
    End If
    quant_band_stereo = cm
End Function

' #ifndef DISABLE_UPDATE_DRAFT
' (VB) norm, norm2: celt_norm_buf offsets
Private Sub special_hybrid_folding(ByVal m As Long, ByVal norm As Long, ByVal norm2 As Long, ByVal start As Long, ByVal M_ As Long, ByVal dual_stereo As Long)
    Dim n1 As Long, n2 As Long
    n1 = M_ * (celt_modes(m).eBands(start + 1) - celt_modes(m).eBands(start))
    n2 = M_ * (celt_modes(m).eBands(start + 2) - celt_modes(m).eBands(start + 1))
    ' Duplicate enough of the first band folding data to be able to fold the second band.
    '   Copies no data for CELT-only mode.
    If n2 - n1 > 0 Then
        CopyMemory celt_norm_buf(norm + n1), celt_norm_buf(norm + 2 * n1 - n2), (n2 - n1) * 4
        If dual_stereo <> 0 Then CopyMemory celt_norm_buf(norm2 + n1), celt_norm_buf(norm2 + 2 * n1 - n2), (n2 - n1) * 4
    End If
End Sub
' #endif

' (VB) X_, Y_ (-1 = NULL): celt_norm_buf offsets; const celt_ener *bandE is the encoder's (NULL here: not a
' parameter); collapse_masks, pulses, tf_res: arrays from 0; opus_uint32 *seed is seed (ByRef)
Public Sub quant_all_bands(ByVal encode As Long, ByVal m As Long, ByVal start As Long, ByVal end_ As Long, _
        ByVal X_ As Long, ByVal Y_ As Long, collapse_masks() As Byte, _
        pulses() As Long, ByVal shortBlocks As Long, ByVal spread As Long, _
        ByVal dual_stereo As Long, ByVal intensity As Long, tf_res() As Long, ByVal total_bits As Long, _
        ByVal balance As Long, ec As ec_ctx, ByVal LM As Long, ByVal codedBands As Long, _
        seed As Long, ByVal complexity As Long, ByVal arch As Long, ByVal disable_inv As Long)
    Dim i As Long, remaining_bits As Long, norm As Long, norm2 As Long, lowband_scratch As Long
    Dim B_ As Long, M_ As Long, lowband_offset As Long, update_lowband As Long, C As Long, norm_offset As Long
    Dim theta_rdo As Long, resynth As Long, ctx As band_ctx
    Dim tell As Long, b As Long, N As Long, curr_balance As Long, effective_lowband As Long
    Dim X As Long, Y As Long, tf_change As Long, x_cm As Long, y_cm As Long, last As Long
    Dim fold_start As Long, fold_end As Long, fold_i As Long, j As Long, lb As Long, lbo As Long, nbEBands As Long
    Dim eBands() As Integer

    eBands = celt_modes(m).eBands
    nbEBands = celt_modes(m).nbEBands
    update_lowband = 1
    If Y_ <> -1 Then C = 2 Else C = 1
    If encode <> 0 And Y_ <> -1 And dual_stereo = 0 And complexity >= 8 Then theta_rdo = 1
    ' #ifdef RESYNTH ... #else
    If encode = 0 Or theta_rdo <> 0 Then resynth = 1
    ' #endif

    M_ = celt_pow2(LM)
    If shortBlocks <> 0 Then B_ = M_ Else B_ = 1
    norm_offset = M_ * eBands(start)
    ' No need to allocate norm for the last band because we don't need an
    '   output in that band.
    ' ALLOC(_norm, C*(M*eBands[m->nbEBands-1]-norm_offset), celt_norm): (VB) celt_norm_buf from NORM_BASE
    norm = NORM_BASE
    norm2 = norm + M_ * eBands(nbEBands - 1) - norm_offset

    ' For decoding, we can use the last band as scratch space because we don't need that
    '   scratch space for the last band and we don't care about the data there until we're
    '   decoding the last band.
    If encode <> 0 And resynth <> 0 Then
        ' (encoder only: resynth_alloc, _lowband_scratch, X_save, Y_save, X_save2, Y_save2, norm_save2)
    Else
        lowband_scratch = X_ + M_ * eBands(celt_modes(m).effEBands - 1)
    End If

    lowband_offset = 0
    ' ctx.bandE = bandE; ctx.ec = ec;
    ctx.encode = encode
    ctx.intensity = intensity
    ctx.m = m
    ctx.seed = seed
    ctx.spread = spread
    ctx.arch = arch
    ctx.disable_inv = disable_inv
    ctx.resynth = resynth
    ctx.theta_round = 0
    ' Avoid injecting noise in the first band on transients.
    If B_ > 1 Then ctx.avoid_split_noise = 1 Else ctx.avoid_split_noise = 0
    For i = start To end_ - 1
        effective_lowband = -1
        tf_change = 0

        ctx.i = i
        If i = end_ - 1 Then last = 1 Else last = 0

        X = X_ + M_ * eBands(i)
        If Y_ <> -1 Then
            Y = Y_ + M_ * eBands(i)
        Else
            Y = -1
        End If
        N = M_ * eBands(i + 1) - M_ * eBands(i)
        ' celt_assert(N > 0);
        tell = ec_tell_frac(ec)

        ' Compute how many bits we want to allocate to this band
        If i <> start Then balance = balance - tell
        remaining_bits = total_bits - tell - 1
        ctx.remaining_bits = remaining_bits
        If i <= codedBands - 1 Then
            If codedBands - i < 3 Then
                curr_balance = balance \ (codedBands - i)       ' celt_sudiv(balance, IMIN(3, codedBands-i))
            Else
                curr_balance = balance \ 3
            End If
            b = pulses(i) + curr_balance
            If remaining_bits + 1 < b Then b = remaining_bits + 1
            If 16383 < b Then b = 16383
            If b < 0 Then b = 0                         ' IMAX(0, IMIN(16383, IMIN(remaining_bits+1,pulses[i]+curr_balance)))
        Else
            b = 0
        End If

        ' #ifndef DISABLE_UPDATE_DRAFT
        If resynth <> 0 And (M_ * eBands(i) - N >= M_ * eBands(start) Or i = start + 1) And (update_lowband <> 0 Or lowband_offset = 0) Then
            lowband_offset = i
        End If
        If i = start + 1 Then special_hybrid_folding m, norm, norm2, start, M_, dual_stereo
        ' #else ... #endif

        tf_change = tf_res(i)
        ctx.tf_change = tf_change
        If i >= celt_modes(m).effEBands Then
            X = norm
            If Y_ <> -1 Then Y = norm
            lowband_scratch = -1
        End If
        If last <> 0 And theta_rdo = 0 Then lowband_scratch = -1

        ' Get a conservative estimate of the collapse_mask's for the bands we're
        '   going to be folding from.
        If lowband_offset <> 0 And (spread <> SPREAD_AGGRESSIVE Or B_ > 1 Or tf_change < 0) Then
            ' This ensures we never repeat spectral content within one band
            effective_lowband = M_ * eBands(lowband_offset) - norm_offset - N
            If effective_lowband < 0 Then effective_lowband = 0
            fold_start = lowband_offset
            Do
                fold_start = fold_start - 1
            Loop While M_ * eBands(fold_start) > effective_lowband + norm_offset
            fold_end = lowband_offset - 1
            ' #ifndef DISABLE_UPDATE_DRAFT
            Do
                fold_end = fold_end + 1
            Loop While fold_end < i And M_ * eBands(fold_end) < effective_lowband + norm_offset + N
            ' #else ... #endif
            x_cm = 0: y_cm = 0
            fold_i = fold_start
            Do
                x_cm = x_cm Or collapse_masks(fold_i * C + 0)
                y_cm = y_cm Or collapse_masks(fold_i * C + C - 1)
                fold_i = fold_i + 1
            Loop While fold_i < fold_end
        ' Otherwise, we'll be using the LCG to fold, so all blocks will (almost
        '   always) be non-zero.
        Else
            x_cm = celt_pow2(B_) - 1
            y_cm = x_cm
        End If

        If dual_stereo <> 0 And i = intensity Then
            ' Switch off dual stereo to do intensity.
            dual_stereo = 0
            If resynth <> 0 Then
                For j = 0 To M_ * eBands(i) - norm_offset - 1
                    celt_norm_buf(norm + j) = 0.5! * (celt_norm_buf(norm + j) + celt_norm_buf(norm2 + j))  ' HALF32(norm[j]+norm2[j])
                Next
            End If
        End If
        If last <> 0 Then lbo = -1 Else lbo = M_ * eBands(i) - norm_offset      ' (VB) last?NULL:norm(2)+lbo
        If dual_stereo <> 0 Then
            If effective_lowband <> -1 Then lb = norm + effective_lowband Else lb = -1
            If lbo <> -1 Then lbo = norm + lbo
            x_cm = quant_band(ctx, X, N, b \ 2, B_, lb, LM, lbo, 1!, lowband_scratch, x_cm, ec)
            If effective_lowband <> -1 Then lb = norm2 + effective_lowband Else lb = -1
            If last <> 0 Then lbo = -1 Else lbo = norm2 + M_ * eBands(i) - norm_offset
            y_cm = quant_band(ctx, Y, N, b \ 2, B_, lb, LM, lbo, 1!, lowband_scratch, y_cm, ec)
        Else
            If effective_lowband <> -1 Then lb = norm + effective_lowband Else lb = -1
            If lbo <> -1 Then lbo = norm + lbo
            If Y <> -1 Then
                If theta_rdo <> 0 And i < intensity Then
                    ' (encoder only: the theta RDO, encoding with theta rounded down and up and keeping the
                    '  better one by compute_channel_weights' distortion)
                Else
                    ctx.theta_round = 0
                    x_cm = quant_band_stereo(ctx, X, Y, N, b, B_, lb, LM, lbo, lowband_scratch, x_cm Or y_cm, ec)
                End If
            Else
                x_cm = quant_band(ctx, X, N, b, B_, lb, LM, lbo, 1!, lowband_scratch, x_cm Or y_cm, ec)
            End If
            y_cm = x_cm
        End If
        collapse_masks(i * C + 0) = x_cm And &HFF       ' (unsigned char)x_cm
        collapse_masks(i * C + C - 1) = y_cm And &HFF
        balance = balance + pulses(i) + tell

        ' Update the folding position only as long as we have 1 bit/sample depth.
        If b > N * 8 Then update_lowband = 1 Else update_lowband = 0    ' b>(N<<BITRES)
        ' We only need to avoid noise on a split for the first band. After that, we
        '   have folding.
        ctx.avoid_split_noise = 0
    Next
    seed = ctx.seed
End Sub
