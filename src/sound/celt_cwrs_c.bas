Attribute VB_Name = "celt_cwrs_c"
' ref/opus/celt (libopus 1.5.2): cwrs.c, the decoder's side, SMALL_FOOTPRINT and CUSTOM_MODES undefined
' (cluster K, fork O2): cwrsi and decode_pulses with the CELT_PVQ_U table (celt_tables_vb's CELT_PVQ_U_DATA,
' generated from the C text). Not ported: icwrs and encode_pulses (the encoder's), log2_frac and
' get_required_bits (CUSTOM_MODES only), the SMALL_FOOTPRINT versions (unext, uprev, ncwrs_urow, icwrs1).
' (VB) An opus_uint32 codeword index is a Double (exact below 2^53); the table holds Doubles. int *_y is
' y_() (Longs) from yo.
'
' Although derived separately, the pulse vector coding scheme is equivalent to
'   a Pyramid Vector Quantizer \cite{Fis86}.
'  (cwrs.c's notes on V(N,K) and U(N,K): see the C file)
Option Explicit

' CELT_PVQ_U_ROW[15]: (VB) the rows' offsets in CELT_PVQ_U_DATA (CUSTOM_MODES undefined)
Private CELT_PVQ_U_ROW(0 To 14) As Long
Private cwrs_rows_done As Boolean

Private Sub cwrs_rows_init()
    CELT_PVQ_U_ROW(0) = 0: CELT_PVQ_U_ROW(1) = 176: CELT_PVQ_U_ROW(2) = 351
    CELT_PVQ_U_ROW(3) = 525: CELT_PVQ_U_ROW(4) = 698: CELT_PVQ_U_ROW(5) = 870
    CELT_PVQ_U_ROW(6) = 1041: CELT_PVQ_U_ROW(7) = 1131: CELT_PVQ_U_ROW(8) = 1178
    CELT_PVQ_U_ROW(9) = 1207: CELT_PVQ_U_ROW(10) = 1226: CELT_PVQ_U_ROW(11) = 1240
    CELT_PVQ_U_ROW(12) = 1248: CELT_PVQ_U_ROW(13) = 1254: CELT_PVQ_U_ROW(14) = 1257
    celt_tables_init
    cwrs_rows_done = True
End Sub

' #if !defined(SMALL_FOOTPRINT)

'U(N,K) = U(K,N) := N>0?K>0?U(N-1,K)+U(N,K-1)+U(N-1,K-1):0:K>0?1:0
' # define CELT_PVQ_U(_n,_k) (CELT_PVQ_U_ROW[IMIN(_n,_k)][IMAX(_n,_k)])
Private Function CELT_PVQ_U(ByVal n As Long, ByVal k As Long) As Double
    If n < k Then
        CELT_PVQ_U = CELT_PVQ_U_DATA(CELT_PVQ_U_ROW(n) + k)
    Else
        CELT_PVQ_U = CELT_PVQ_U_DATA(CELT_PVQ_U_ROW(k) + n)
    End If
End Function

'V(N,K) := U(N,K)+U(N,K+1) = the number of PVQ codewords for a band of size N
'   with K pulses allocated to it.
' # define CELT_PVQ_V(_n,_k) (CELT_PVQ_U(_n,_k)+CELT_PVQ_U(_n,(_k)+1))
Private Function CELT_PVQ_V(ByVal n As Long, ByVal k As Long) As Double
    CELT_PVQ_V = CELT_PVQ_U(n, k) + CELT_PVQ_U(n, k + 1)
End Function

' (VB) the result: opus_val32 yy (the float build's: a float)
Private Function cwrsi(ByVal n_ As Long, ByVal k_ As Long, ByVal i_ As Double, y_() As Long, ByVal yo As Long) As Single
    Dim p As Double, s As Long, k0 As Long, val As Long, yy As Single, q As Double, row As Long
    ' celt_assert(_k>0);
    ' celt_assert(_n>1);
    Do While n_ > 2
        'Lots of pulses case:
        If k_ >= n_ Then
            row = CELT_PVQ_U_ROW(n_)
            'Are the pulses in this dimension negative?
            p = CELT_PVQ_U_DATA(row + k_ + 1)
            ' s=-(_i>=p); _i-=p&s;
            If i_ >= p Then
                s = -1
                i_ = i_ - p
            Else
                s = 0
            End If
            'Count how many pulses were placed in this dimension.
            k0 = k_
            q = CELT_PVQ_U_DATA(row + n_)
            If q > i_ Then
                ' celt_sig_assert(p>q);
                k_ = n_
                Do
                    k_ = k_ - 1
                    p = CELT_PVQ_U_DATA(CELT_PVQ_U_ROW(k_) + n_)
                Loop While p > i_
            Else
                p = CELT_PVQ_U_DATA(row + k_)
                Do While p > i_
                    k_ = k_ - 1
                    p = CELT_PVQ_U_DATA(row + k_)
                Loop
            End If
            i_ = i_ - p
            val = (k0 - k_ + s) Xor s
            y_(yo) = val
            yo = yo + 1
            yy = yy + CSng(val) * val                   ' MAC16_16(yy,val,val)
        'Lots of dimensions case:
        Else
            'Are there any pulses in this dimension at all?
            p = CELT_PVQ_U_DATA(CELT_PVQ_U_ROW(k_) + n_)
            q = CELT_PVQ_U_DATA(CELT_PVQ_U_ROW(k_ + 1) + n_)
            If p <= i_ And i_ < q Then
                i_ = i_ - p
                y_(yo) = 0
                yo = yo + 1
            Else
                'Are the pulses in this dimension negative?
                If i_ >= q Then
                    s = -1
                    i_ = i_ - q
                Else
                    s = 0
                End If
                'Count how many pulses were placed in this dimension.
                k0 = k_
                Do
                    k_ = k_ - 1
                    p = CELT_PVQ_U_DATA(CELT_PVQ_U_ROW(k_) + n_)
                Loop While p > i_
                i_ = i_ - p
                val = (k0 - k_ + s) Xor s
                y_(yo) = val
                yo = yo + 1
                yy = yy + CSng(val) * val
            End If
        End If
        n_ = n_ - 1
    Loop
    '_n==2
    p = 2 * k_ + 1
    If i_ >= p Then
        s = -1
        i_ = i_ - p
    Else
        s = 0
    End If
    k0 = k_
    k_ = Int((i_ + 1) / 2)                              ' (_i+1)>>1
    If k_ <> 0 Then i_ = i_ - (2 * k_ - 1)
    val = (k0 - k_ + s) Xor s
    y_(yo) = val
    yo = yo + 1
    yy = yy + CSng(val) * val
    '_n==1
    s = -i_                                             ' s=-(int)_i;
    val = (k_ + s) Xor s
    y_(yo) = val
    yy = yy + CSng(val) * val
    cwrsi = yy
End Function

' (VB) int *_y is y_() from yo; the result is opus_val32 (a float)
Public Function decode_pulses(y_() As Long, ByVal yo As Long, ByVal n_ As Long, ByVal k_ As Long, dec As ec_ctx) As Single
    Dim v As Double, i_ As Double
    If Not cwrs_rows_done Then cwrs_rows_init
    v = CELT_PVQ_V(n_, k_)
    If v >= 2147483648# Then v = v - 4294967296#       ' (VB) the opus_uint32 as a Long's bits
    i_ = ec_dec_uint(dec, v)
    If i_ < 0 Then i_ = i_ + 4294967296#
    decode_pulses = cwrsi(n_, k_, i_, y_, yo)
End Function

' #else /* SMALL_FOOTPRINT */ (not compiled)
