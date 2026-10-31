Attribute VB_Name = "celt_laplace_c"
' ref/opus/celt (libopus 1.5.2): laplace.c, the decoder's side (cluster K, fork O2): ec_laplace_decode and
' ec_laplace_get_freq1. Not ported: ec_laplace_encode and ec_laplace_encode_p0 (the encoder's),
' ec_laplace_decode_p0 (DRED's only: ENABLE_DRED is undefined), the #if 0 test main.
Option Explicit

' The minimum probability of an energy delta (out of 32768).
Private Const LAPLACE_LOG_MINP As Long = 0
Private Const LAPLACE_MINP As Long = 1              ' (1<<LAPLACE_LOG_MINP)
' The minimum number of guaranteed representable energy deltas (in one
'    direction).
Private Const LAPLACE_NMIN As Long = 16

' When called, decay is positive and at most 11456.
Private Function ec_laplace_get_freq1(ByVal fs0 As Long, ByVal decay As Long) As Long
    Dim ft As Long
    ft = 32768 - LAPLACE_MINP * (2 * LAPLACE_NMIN) - fs0
    ec_laplace_get_freq1 = (ft * (16384 - decay)) \ 32768      ' >>15 (not negative)
End Function

Public Function ec_laplace_decode(dec As ec_ctx, ByVal fs As Long, ByVal decay As Long) As Long
    Dim val As Long, fl As Long, fm As Long, di As Long, fh As Long
    val = 0
    fm = ec_decode_bin(dec, 15)
    fl = 0
    If fm >= fs Then
        val = val + 1
        fl = fs
        fs = ec_laplace_get_freq1(fs, decay) + LAPLACE_MINP
        ' Search the decaying part of the PDF.
        Do While fs > LAPLACE_MINP And fm >= fl + 2 * fs
            fs = fs * 2
            fl = fl + fs
            fs = ((fs - 2 * LAPLACE_MINP) * decay) \ 32768             ' >>15 (not negative)
            fs = fs + LAPLACE_MINP
            val = val + 1
        Loop
        ' Everything beyond that has probability LAPLACE_MINP.
        If fs <= LAPLACE_MINP Then
            di = (fm - fl) \ 2                                      ' >>(LAPLACE_LOG_MINP+1)
            val = val + di
            fl = fl + 2 * di * LAPLACE_MINP
        End If
        If fm < fl + fs Then
            val = -val
        Else
            fl = fl + fs
        End If
    End If
    ' celt_assert(fl<32768);
    ' celt_assert(fs>0);
    ' celt_assert(fl<=fm);
    ' celt_assert(fm<IMIN(fl+fs,32768));
    fh = fl + fs
    If fh > 32768 Then fh = 32768                                   ' IMIN(fl+fs,32768)
    ec_dec_update dec, fl, fh, 32768
    ec_laplace_decode = val
End Function
