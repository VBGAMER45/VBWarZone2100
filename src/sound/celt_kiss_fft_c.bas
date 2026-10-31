Attribute VB_Name = "celt_kiss_fft_c"
' ref/opus/celt (libopus 1.5.2): kiss_fft.c with _kiss_fft_guts.h's float macros, the float build without
' CUSTOM_MODES or SIMD (kiss_fft_scalar = float) (cluster K, fork O2). The decoder reaches opus_fft_impl only
' (clt_mdct_backward_c), on the static mode's precomputed states (celt_modes_c's celt_fft_states). Not ported:
' opus_fft_c and opus_ifft_c (the encoder's clt_mdct_forward and the tests'), and the CUSTOM_MODES part
' (compute_bitrev_table, kf_factor, compute_twiddles, opus_fft_alloc*, opus_fft_free*) with kf_bfly2's m==1
' branch, which only custom modes reach.
' (VB) A kiss_fft_cpx *Fout is a Single array and an index: complex k is fout(fo + 2*k) (.r) and
' fout(fo + 2*k + 1) (.i). The twiddles are passed as their .r and .i arrays.
'
'  Copyright (c) 2003-2004, Mark Borgerding
'  Lots of modifications by Jean-Marc Valin
'  Copyright (c) 2005-2007, Xiph.Org Foundation
'  Copyright (c) 2008,      Xiph.Org Foundation, CSIRO
'
'  All rights reserved.
'  (the BSD licence terms of kiss_fft.c)
'
' This code is originally from Mark Borgerding's KISS-FFT but has been
'   heavily modified to better suit Opus
Option Explicit

' The guts header contains all the multiplication and addition macros that are defined for
'   complex numbers.  It also delares the kf_ internal functions.

Private Sub kf_bfly2(fout() As Single, ByVal fo As Long, ByVal m As Long, ByVal N As Long)
    Dim f As Long, f2 As Long, i As Long
    Dim tw As Single, tr As Single, ti As Single
    ' #ifdef CUSTOM_MODES: the m==1 case (not compiled)
    tw = 0.7071068!                                     ' QCONST16(0.7071067812f, 15)
    ' We know that m==4 here because the radix-2 is just after a radix-4
    ' celt_assert(m==4);
    f = fo
    For i = 0 To N - 1
        f2 = f + 8                                      ' Fout2 = Fout + 4
        ' t = Fout2[0]; C_SUB( Fout2[0] ,  Fout[0] , t ); C_ADDTO( Fout[0] ,  t );
        tr = fout(f2): ti = fout(f2 + 1)
        fout(f2) = fout(f) - tr: fout(f2 + 1) = fout(f + 1) - ti
        fout(f) = fout(f) + tr: fout(f + 1) = fout(f + 1) + ti

        tr = (fout(f2 + 2) + fout(f2 + 3)) * tw         ' t.r = S_MUL(ADD32_ovflw(Fout2[1].r, Fout2[1].i), tw);
        ti = (fout(f2 + 3) - fout(f2 + 2)) * tw         ' t.i = S_MUL(SUB32_ovflw(Fout2[1].i, Fout2[1].r), tw);
        fout(f2 + 2) = fout(f + 2) - tr: fout(f2 + 3) = fout(f + 3) - ti
        fout(f + 2) = fout(f + 2) + tr: fout(f + 3) = fout(f + 3) + ti

        tr = fout(f2 + 5)                               ' t.r = Fout2[2].i;
        ti = -fout(f2 + 4)                              ' t.i = -Fout2[2].r;
        fout(f2 + 4) = fout(f + 4) - tr: fout(f2 + 5) = fout(f + 5) - ti
        fout(f + 4) = fout(f + 4) + tr: fout(f + 5) = fout(f + 5) + ti

        tr = (fout(f2 + 7) - fout(f2 + 6)) * tw         ' t.r = S_MUL(SUB32_ovflw(Fout2[3].i, Fout2[3].r), tw);
        ti = -(fout(f2 + 7) + fout(f2 + 6)) * tw        ' t.i = S_MUL(NEG32_ovflw(ADD32_ovflw(Fout2[3].i, Fout2[3].r)), tw);
        fout(f2 + 6) = fout(f + 6) - tr: fout(f2 + 7) = fout(f + 7) - ti
        fout(f + 6) = fout(f + 6) + tr: fout(f + 7) = fout(f + 7) + ti
        f = f + 16                                      ' Fout += 8
    Next
End Sub

Private Sub kf_bfly4(fout() As Single, ByVal fo As Long, ByVal fstride As Long, tw_r() As Single, tw_i() As Single, ByVal m As Long, ByVal N As Long, ByVal mm As Long)
    Dim i As Long, j As Long, f As Long, om As Long, om2 As Long, om3 As Long
    Dim tw1 As Long, tw2 As Long, tw3 As Long, fstride2 As Long, fstride3 As Long
    Dim s0r As Single, s0i As Single, s1r As Single, s1i As Single, s2r As Single, s2i As Single
    Dim s3r As Single, s3i As Single, s4r As Single, s4i As Single, s5r As Single, s5i As Single
    Dim ar As Single, ai As Single, fr As Single, fi As Single

    If m = 1 Then
        ' Degenerate case where all the twiddles are 1.
        f = fo
        For i = 0 To N - 1
            fr = fout(f): fi = fout(f + 1)
            s0r = fr - fout(f + 4): s0i = fi - fout(f + 5)                  ' C_SUB( scratch0 , *Fout, Fout[2] );
            fr = fr + fout(f + 4): fi = fi + fout(f + 5)                    ' C_ADDTO(*Fout, Fout[2]);
            s1r = fout(f + 2) + fout(f + 6): s1i = fout(f + 3) + fout(f + 7) ' C_ADD( scratch1 , Fout[1] , Fout[3] );
            fout(f + 4) = fr - s1r: fout(f + 5) = fi - s1i                  ' C_SUB( Fout[2], *Fout, scratch1 );
            fout(f) = fr + s1r: fout(f + 1) = fi + s1i                      ' C_ADDTO( *Fout , scratch1 );
            s1r = fout(f + 2) - fout(f + 6): s1i = fout(f + 3) - fout(f + 7) ' C_SUB( scratch1 , Fout[1] , Fout[3] );

            fout(f + 2) = s0r + s1i                     ' Fout[1].r = ADD32_ovflw(scratch0.r, scratch1.i);
            fout(f + 3) = s0i - s1r                     ' Fout[1].i = SUB32_ovflw(scratch0.i, scratch1.r);
            fout(f + 6) = s0r - s1i                     ' Fout[3].r = SUB32_ovflw(scratch0.r, scratch1.i);
            fout(f + 7) = s0i + s1r                     ' Fout[3].i = ADD32_ovflw(scratch0.i, scratch1.r);
            f = f + 8                                   ' Fout+=4;
        Next
    Else
        om = 2 * m                                      ' (VB) Fout[m], Fout[m2], Fout[m3] as offsets
        om2 = 4 * m                                     ' m2=2*m
        om3 = 6 * m                                     ' m3=3*m
        fstride2 = fstride * 2
        fstride3 = fstride * 3
        For i = 0 To N - 1
            f = fo + 2 * i * mm                         ' Fout = Fout_beg + i*mm;
            tw1 = 0: tw2 = 0: tw3 = 0                   ' tw3 = tw2 = tw1 = st->twiddles;
            ' m is guaranteed to be a multiple of 4.
            For j = 0 To m - 1
                ar = fout(f + om): ai = fout(f + om + 1)
                s0r = ar * tw_r(tw1) - ai * tw_i(tw1)   ' C_MUL(scratch[0],Fout[m] , *tw1 );
                s0i = ar * tw_i(tw1) + ai * tw_r(tw1)
                ar = fout(f + om2): ai = fout(f + om2 + 1)
                s1r = ar * tw_r(tw2) - ai * tw_i(tw2)   ' C_MUL(scratch[1],Fout[m2] , *tw2 );
                s1i = ar * tw_i(tw2) + ai * tw_r(tw2)
                ar = fout(f + om3): ai = fout(f + om3 + 1)
                s2r = ar * tw_r(tw3) - ai * tw_i(tw3)   ' C_MUL(scratch[2],Fout[m3] , *tw3 );
                s2i = ar * tw_i(tw3) + ai * tw_r(tw3)

                fr = fout(f): fi = fout(f + 1)
                s5r = fr - s1r: s5i = fi - s1i          ' C_SUB( scratch[5] , *Fout, scratch[1] );
                fr = fr + s1r: fi = fi + s1i            ' C_ADDTO(*Fout, scratch[1]);
                s3r = s0r + s2r: s3i = s0i + s2i        ' C_ADD( scratch[3] , scratch[0] , scratch[2] );
                s4r = s0r - s2r: s4i = s0i - s2i        ' C_SUB( scratch[4] , scratch[0] , scratch[2] );
                fout(f + om2) = fr - s3r                ' C_SUB( Fout[m2], *Fout, scratch[3] );
                fout(f + om2 + 1) = fi - s3i
                tw1 = tw1 + fstride
                tw2 = tw2 + fstride2
                tw3 = tw3 + fstride3
                fout(f) = fr + s3r                      ' C_ADDTO( *Fout , scratch[3] );
                fout(f + 1) = fi + s3i

                fout(f + om) = s5r + s4i                ' Fout[m].r = ADD32_ovflw(scratch[5].r, scratch[4].i);
                fout(f + om + 1) = s5i - s4r            ' Fout[m].i = SUB32_ovflw(scratch[5].i, scratch[4].r);
                fout(f + om3) = s5r - s4i               ' Fout[m3].r = SUB32_ovflw(scratch[5].r, scratch[4].i);
                fout(f + om3 + 1) = s5i + s4r           ' Fout[m3].i = ADD32_ovflw(scratch[5].i, scratch[4].r);
                f = f + 2                               ' ++Fout;
            Next
        Next
    End If
End Sub

' #ifndef RADIX_TWO_ONLY

Private Sub kf_bfly3(fout() As Single, ByVal fo As Long, ByVal fstride As Long, tw_r() As Single, tw_i() As Single, ByVal m As Long, ByVal N As Long, ByVal mm As Long)
    Dim i As Long, k As Long, f As Long, om As Long, om2 As Long, tw1 As Long, tw2 As Long, fstride2 As Long
    Dim s0r As Single, s0i As Single, s1r As Single, s1i As Single, s2r As Single, s2i As Single
    Dim s3r As Single, s3i As Single, ar As Single, ai As Single, epi3i As Single, mr As Single, mi As Single

    om = 2 * m
    om2 = 4 * m                                         ' m2 = 2*m
    fstride2 = fstride * 2
    ' #ifdef FIXED_POINT ... #else
    epi3i = tw_i(fstride * m)                           ' epi3 = st->twiddles[fstride*m]; (only .i is used)
    For i = 0 To N - 1
        f = fo + 2 * i * mm                             ' Fout = Fout_beg + i*mm;
        tw1 = 0: tw2 = 0                                ' tw1=tw2=st->twiddles;
        ' For non-custom modes, m is guaranteed to be a multiple of 4.
        k = m
        Do
            ar = fout(f + om): ai = fout(f + om + 1)
            s1r = ar * tw_r(tw1) - ai * tw_i(tw1)       ' C_MUL(scratch[1],Fout[m] , *tw1);
            s1i = ar * tw_i(tw1) + ai * tw_r(tw1)
            ar = fout(f + om2): ai = fout(f + om2 + 1)
            s2r = ar * tw_r(tw2) - ai * tw_i(tw2)       ' C_MUL(scratch[2],Fout[m2] , *tw2);
            s2i = ar * tw_i(tw2) + ai * tw_r(tw2)

            s3r = s1r + s2r: s3i = s1i + s2i            ' C_ADD(scratch[3],scratch[1],scratch[2]);
            s0r = s1r - s2r: s0i = s1i - s2i            ' C_SUB(scratch[0],scratch[1],scratch[2]);
            tw1 = tw1 + fstride
            tw2 = tw2 + fstride2

            mr = fout(f) - s3r * 0.5!                   ' Fout[m].r = SUB32_ovflw(Fout->r, HALF_OF(scratch[3].r));
            mi = fout(f + 1) - s3i * 0.5!               ' Fout[m].i = SUB32_ovflw(Fout->i, HALF_OF(scratch[3].i));

            s0r = s0r * epi3i                           ' C_MULBYSCALAR( scratch[0] , epi3.i );
            s0i = s0i * epi3i

            fout(f) = fout(f) + s3r                     ' C_ADDTO(*Fout,scratch[3]);
            fout(f + 1) = fout(f + 1) + s3i

            fout(f + om2) = mr + s0i                    ' Fout[m2].r = ADD32_ovflw(Fout[m].r, scratch[0].i);
            fout(f + om2 + 1) = mi - s0r                ' Fout[m2].i = SUB32_ovflw(Fout[m].i, scratch[0].r);

            fout(f + om) = mr - s0i                     ' Fout[m].r = SUB32_ovflw(Fout[m].r, scratch[0].i);
            fout(f + om + 1) = mi + s0r                 ' Fout[m].i = ADD32_ovflw(Fout[m].i, scratch[0].r);

            f = f + 2                                   ' ++Fout;
            k = k - 1
        Loop While k <> 0
    Next
End Sub

' #ifndef OVERRIDE_kf_bfly5
Private Sub kf_bfly5(fout() As Single, ByVal fo As Long, ByVal fstride As Long, tw_r() As Single, tw_i() As Single, ByVal m As Long, ByVal N As Long, ByVal mm As Long)
    Dim f0 As Long, f1 As Long, f2 As Long, f3 As Long, f4 As Long, i As Long, u As Long
    Dim t1 As Long, t2 As Long, t3 As Long, t4 As Long
    Dim yar As Single, yai As Single, ybr As Single, ybi As Single
    Dim s0r As Single, s0i As Single, s1r As Single, s1i As Single, s2r As Single, s2i As Single
    Dim s3r As Single, s3i As Single, s4r As Single, s4i As Single, s5r As Single, s5i As Single
    Dim s6r As Single, s6i As Single, s7r As Single, s7i As Single, s8r As Single, s8i As Single
    Dim s9r As Single, s9i As Single, s10r As Single, s10i As Single, s11r As Single, s11i As Single
    Dim s12r As Single, s12i As Single, ar As Single, ai As Single

    ' #ifdef FIXED_POINT ... #else
    yar = tw_r(fstride * m): yai = tw_i(fstride * m)            ' ya = st->twiddles[fstride*m];
    ybr = tw_r(fstride * 2 * m): ybi = tw_i(fstride * 2 * m)    ' yb = st->twiddles[fstride*2*m];
    ' tw=st->twiddles;

    For i = 0 To N - 1
        f0 = fo + 2 * i * mm                            ' Fout = Fout_beg + i*mm; Fout0=Fout;
        f1 = f0 + 2 * m                                 ' Fout1=Fout0+m;
        f2 = f0 + 4 * m                                 ' Fout2=Fout0+2*m;
        f3 = f0 + 6 * m                                 ' Fout3=Fout0+3*m;
        f4 = f0 + 8 * m                                 ' Fout4=Fout0+4*m;

        ' For non-custom modes, m is guaranteed to be a multiple of 4.
        For u = 0 To m - 1
            s0r = fout(f0): s0i = fout(f0 + 1)         ' scratch[0] = *Fout0;

            t1 = u * fstride: t2 = 2 * t1: t3 = 3 * t1: t4 = 4 * t1
            ar = fout(f1): ai = fout(f1 + 1)
            s1r = ar * tw_r(t1) - ai * tw_i(t1)         ' C_MUL(scratch[1] ,*Fout1, tw[u*fstride]);
            s1i = ar * tw_i(t1) + ai * tw_r(t1)
            ar = fout(f2): ai = fout(f2 + 1)
            s2r = ar * tw_r(t2) - ai * tw_i(t2)         ' C_MUL(scratch[2] ,*Fout2, tw[2*u*fstride]);
            s2i = ar * tw_i(t2) + ai * tw_r(t2)
            ar = fout(f3): ai = fout(f3 + 1)
            s3r = ar * tw_r(t3) - ai * tw_i(t3)         ' C_MUL(scratch[3] ,*Fout3, tw[3*u*fstride]);
            s3i = ar * tw_i(t3) + ai * tw_r(t3)
            ar = fout(f4): ai = fout(f4 + 1)
            s4r = ar * tw_r(t4) - ai * tw_i(t4)         ' C_MUL(scratch[4] ,*Fout4, tw[4*u*fstride]);
            s4i = ar * tw_i(t4) + ai * tw_r(t4)

            s7r = s1r + s4r: s7i = s1i + s4i            ' C_ADD( scratch[7],scratch[1],scratch[4]);
            s10r = s1r - s4r: s10i = s1i - s4i          ' C_SUB( scratch[10],scratch[1],scratch[4]);
            s8r = s2r + s3r: s8i = s2i + s3i            ' C_ADD( scratch[8],scratch[2],scratch[3]);
            s9r = s2r - s3r: s9i = s2i - s3i            ' C_SUB( scratch[9],scratch[2],scratch[3]);

            fout(f0) = s0r + (s7r + s8r)                ' Fout0->r = ADD32_ovflw(Fout0->r, ADD32_ovflw(scratch[7].r, scratch[8].r));
            fout(f0 + 1) = s0i + (s7i + s8i)            ' Fout0->i = ADD32_ovflw(Fout0->i, ADD32_ovflw(scratch[7].i, scratch[8].i));

            s5r = s0r + (s7r * yar + s8r * ybr)         ' scratch[5].r = ... S_MUL(scratch[7].r,ya.r), S_MUL(scratch[8].r,yb.r)
            s5i = s0i + (s7i * yar + s8i * ybr)

            s6r = s10i * yai + s9i * ybi                ' scratch[6].r =  ADD32_ovflw(S_MUL(scratch[10].i,ya.i), S_MUL(scratch[9].i,yb.i));
            s6i = -(s10r * yai + s9r * ybi)             ' scratch[6].i = NEG32_ovflw(ADD32_ovflw(S_MUL(scratch[10].r,ya.i), S_MUL(scratch[9].r,yb.i)));

            fout(f1) = s5r - s6r: fout(f1 + 1) = s5i - s6i      ' C_SUB(*Fout1,scratch[5],scratch[6]);
            fout(f4) = s5r + s6r: fout(f4 + 1) = s5i + s6i      ' C_ADD(*Fout4,scratch[5],scratch[6]);

            s11r = s0r + (s7r * ybr + s8r * yar)        ' scratch[11].r = ... S_MUL(scratch[7].r,yb.r), S_MUL(scratch[8].r,ya.r)
            s11i = s0i + (s7i * ybr + s8i * yar)
            s12r = s9i * yai - s10i * ybi               ' scratch[12].r = SUB32_ovflw(S_MUL(scratch[9].i,ya.i), S_MUL(scratch[10].i,yb.i));
            s12i = s10r * ybi - s9r * yai               ' scratch[12].i = SUB32_ovflw(S_MUL(scratch[10].r,yb.i), S_MUL(scratch[9].r,ya.i));

            fout(f2) = s11r + s12r: fout(f2 + 1) = s11i + s12i  ' C_ADD(*Fout2,scratch[11],scratch[12]);
            fout(f3) = s11r - s12r: fout(f3 + 1) = s11i - s12i  ' C_SUB(*Fout3,scratch[11],scratch[12]);

            f0 = f0 + 2: f1 = f1 + 2: f2 = f2 + 2: f3 = f3 + 2: f4 = f4 + 2     ' ++Fout0;++Fout1;++Fout2;++Fout3;++Fout4;
        Next
    Next
End Sub
' #endif /* OVERRIDE_kf_bfly5 */

' #endif

' (VB) const kiss_fft_state *st is the state itself; kiss_fft_cpx *fout is fout() from fo
Public Sub opus_fft_impl(st As kiss_fft_state, fout() As Single, ByVal fo As Long)
    Dim m2 As Long, m As Long, p As Long, L As Long, i As Long, shift As Long
    Dim fstride(0 To MAXFACTORS - 1) As Long

    ' st->shift can be -1
    If st.shift > 0 Then shift = st.shift Else shift = 0

    fstride(0) = 1
    L = 0
    Do
        p = st.factors(2 * L)
        m = st.factors(2 * L + 1)
        fstride(L + 1) = fstride(L) * p
        L = L + 1
    Loop While m <> 1
    m = st.factors(2 * L - 1)
    For i = L - 1 To 0 Step -1
        If i <> 0 Then
            m2 = st.factors(2 * i - 1)
        Else
            m2 = 1
        End If
        Select Case st.factors(2 * i)
        Case 2
            kf_bfly2 fout, fo, m, fstride(i)
        Case 4
            kf_bfly4 fout, fo, fstride(i) * celt_pow2(shift), st.twiddles_r, st.twiddles_i, m, fstride(i), m2
        ' #ifndef RADIX_TWO_ONLY
        Case 3
            kf_bfly3 fout, fo, fstride(i) * celt_pow2(shift), st.twiddles_r, st.twiddles_i, m, fstride(i), m2
        Case 5
            kf_bfly5 fout, fo, fstride(i) * celt_pow2(shift), st.twiddles_r, st.twiddles_i, m, fstride(i), m2
        ' #endif
        End Select
        m = m2
    Next
End Sub
