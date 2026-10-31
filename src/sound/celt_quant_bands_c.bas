Attribute VB_Name = "celt_quant_bands_c"
' ref/opus/celt (libopus 1.5.2): quant_bands.c and quant_bands.h, the float build's decoder side (cluster K,
' fork O2): unquant_coarse_energy, unquant_fine_energy, unquant_energy_finalise and their tables (eMeans and
' e_prob_model are celt_tables_vb's, generated from the C text). Not ported (the encoder's): loss_distortion,
' quant_coarse_energy_impl, quant_coarse_energy, quant_fine_energy, quant_energy_finalise, amp2Log2.
' (VB) const CELTMode *m is the mode's handle (celt_modes_c); opus_val16 *oldEBands is oldEBands() from 0.
Option Explicit

' Mean energy in each band quantized in Q4 and converted back to float: eMeans (celt_tables_vb)

' prediction coefficients: 0.9, 0.8, 0.65, 0.5
' static const opus_val16 pred_coef[4] = {29440/32768., 26112/32768., 21248/32768., 16384/32768.};
' static const opus_val16 beta_coef[4] = {30147/32768., 22282/32768., 12124/32768., 6554/32768.};
' static const opus_val16 beta_intra = 4915/32768.;
Private pred_coef(0 To 3) As Single
Private beta_coef(0 To 3) As Single
Private Const beta_intra As Single = 4915! / 32768!

'Parameters of the Laplace-like probability models used for the coarse energy.
'  There is one pair of parameters for each frame size, prediction type
'   (inter/intra), and band number.
'  The first number of each pair is the probability of 0, and the second is the
'   decay rate, both in Q8 precision.
' static const unsigned char e_prob_model[4][2][42]: celt_tables_vb's e_prob_model (flattened)

' static const unsigned char small_energy_icdf[3]={2,1,0};
Private small_energy_icdf(0 To 2) As Byte

Private quant_bands_done As Boolean

Private Sub quant_bands_init()
    pred_coef(0) = 29440! / 32768!: pred_coef(1) = 26112! / 32768!: pred_coef(2) = 21248! / 32768!: pred_coef(3) = 16384! / 32768!
    beta_coef(0) = 30147! / 32768!: beta_coef(1) = 22282! / 32768!: beta_coef(2) = 12124! / 32768!: beta_coef(3) = 6554! / 32768!
    small_energy_icdf(0) = 2: small_energy_icdf(1) = 1: small_energy_icdf(2) = 0
    celt_tables_init
    quant_bands_done = True
End Sub

Public Sub unquant_coarse_energy(ByVal m As Long, ByVal start As Long, ByVal end_ As Long, oldEBands() As Single, ByVal intra As Long, dec As ec_ctx, ByVal C As Long, ByVal LM As Long)
    Dim prob_model As Long, i As Long, c_ As Long, nbEBands As Long
    Dim prev(0 To 1) As Single, coef As Single, beta As Single
    Dim budget As Long, tell As Long, qi As Long, q As Single, tmp As Single, pi_ As Long, k As Long
    If Not quant_bands_done Then quant_bands_init
    prob_model = LM * 84 + intra * 42                   ' e_prob_model[LM][intra]
    nbEBands = celt_modes(m).nbEBands

    If intra <> 0 Then
        coef = 0
        beta = beta_intra
    Else
        beta = beta_coef(LM)
        coef = pred_coef(LM)
    End If

    budget = dec.storage * 8

    ' Decode at a fixed coarse resolution
    For i = start To end_ - 1
        c_ = 0
        Do
            ' It would be better to express this invariant as a
            '   test on C at function entry, but that isn't enough
            '   to make the static analyzer happy.
            ' celt_sig_assert(c<2);
            tell = ec_tell(dec)
            If budget - tell >= 15 Then
                If i < 20 Then pi_ = 2 * i Else pi_ = 40    ' 2*IMIN(i,20)
                qi = ec_laplace_decode(dec, CLng(e_prob_model(prob_model + pi_)) * 128, CLng(e_prob_model(prob_model + pi_ + 1)) * 64)
            ElseIf budget - tell >= 2 Then
                qi = ec_dec_icdf(dec, small_energy_icdf, 0, 2)
                qi = (qi \ 2) Xor -(qi And 1)
            ElseIf budget - tell >= 1 Then
                qi = -ec_dec_bit_logp(dec, 1)
            Else
                qi = -1
            End If
            q = qi                                      ' (opus_val32)SHL32(EXTEND32(qi),DB_SHIFT)

            k = i + c_ * nbEBands
            If oldEBands(k) < -9! Then oldEBands(k) = -9!   ' MAX16(-QCONST16(9.f,DB_SHIFT), oldEBands[...])
            tmp = coef * oldEBands(k) + prev(c_) + q    ' PSHR32(MULT16_16(coef,oldEBands[...]),8) + prev[c] + SHL32(q,7)
            ' #ifdef FIXED_POINT ... #endif
            oldEBands(k) = tmp                          ' PSHR32(tmp, 7)
            prev(c_) = prev(c_) + q - beta * q          ' prev[c] + SHL32(q,7) - MULT16_16(beta,PSHR32(q,8))
            c_ = c_ + 1
        Loop While c_ < C
    Next
End Sub

Public Sub unquant_fine_energy(ByVal m As Long, ByVal start As Long, ByVal end_ As Long, oldEBands() As Single, fine_quant() As Long, dec As ec_ctx, ByVal C As Long)
    Dim i As Long, c_ As Long, q2 As Long, offset As Single, nbEBands As Long
    nbEBands = celt_modes(m).nbEBands
    ' Decode finer resolution
    For i = start To end_ - 1
        If fine_quant(i) <= 0 Then GoTo continue_
        c_ = 0
        Do
            q2 = ec_dec_bits(dec, fine_quant(i))
            ' #ifdef FIXED_POINT ... #else
            offset = (q2 + 0.5!) * CSng(celt_pow2(14 - fine_quant(i))) * (1! / 16384) - 0.5!
            oldEBands(i + c_ * nbEBands) = oldEBands(i + c_ * nbEBands) + offset
            c_ = c_ + 1
        Loop While c_ < C
continue_:
    Next
End Sub

Public Sub unquant_energy_finalise(ByVal m As Long, ByVal start As Long, ByVal end_ As Long, oldEBands() As Single, fine_quant() As Long, fine_priority() As Long, ByVal bits_left As Long, dec As ec_ctx, ByVal C As Long)
    Dim i As Long, prio As Long, c_ As Long, q2 As Long, offset As Single, nbEBands As Long
    nbEBands = celt_modes(m).nbEBands
    ' Use up the remaining bits
    For prio = 0 To 1
        i = start
        Do While i < end_ And bits_left >= C
            If fine_quant(i) >= MAX_FINE_BITS Or fine_priority(i) <> prio Then GoTo continue_
            c_ = 0
            Do
                q2 = ec_dec_bits(dec, 1)
                ' #ifdef FIXED_POINT ... #else
                offset = (q2 - 0.5!) * CSng(celt_pow2(14 - fine_quant(i) - 1)) * (1! / 16384)
                oldEBands(i + c_ * nbEBands) = oldEBands(i + c_ * nbEBands) + offset
                bits_left = bits_left - 1
                c_ = c_ + 1
            Loop While c_ < C
continue_:
            i = i + 1
        Loop
    Next
End Sub
