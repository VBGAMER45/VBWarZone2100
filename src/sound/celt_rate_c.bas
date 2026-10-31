Attribute VB_Name = "celt_rate_c"
' ref/opus/celt (libopus 1.5.2): rate.c and rate.h (its inlines get_pulses, bits2pulses, pulses2bits), the
' decoder's side (cluster K, fork O2). The encode branches of interp_bits2pulses are comments. Not ported:
' fits_in32 and compute_pulse_cache (CUSTOM_MODES only).
' (VB) const CELTMode *m is the mode's handle (celt_modes_c); int arrays are Long arrays from 0 (one entry per
' band); ec_ctx *ec is ec (ByRef). celt_udiv of the C here only divides non-negative values: `\`.
Option Explicit

Public Const MAX_PSEUDO As Long = 40
Public Const LOG_MAX_PSEUDO As Long = 6

Public Const CELT_MAX_PULSES As Long = 128

Public Const MAX_FINE_BITS As Long = 8

Public Const FINE_OFFSET As Long = 21
Public Const QTHETA_OFFSET As Long = 4
Public Const QTHETA_OFFSET_TWOPHASE As Long = 16

Private Const ALLOC_STEPS As Long = 6

' static const unsigned char LOG2_FRAC_TABLE[24]={
'    0,
'    8,13,
'   16,19,21,23,
'   24,26,27,28,29,30,31,32,
'   32,33,34,34,35,36,36,37,37
' };
Private LOG2_FRAC_TABLE(0 To 23) As Long
Private rate_done As Boolean

Private Sub rate_init()
    Dim v As Variant, i As Long
    v = Array(0, 8, 13, 16, 19, 21, 23, 24, 26, 27, 28, 29, 30, 31, 32, 32, 33, 34, 34, 35, 36, 36, 37, 37)
    For i = 0 To 23
        LOG2_FRAC_TABLE(i) = v(i)
    Next
    rate_done = True
End Sub

' rate.h

Public Function get_pulses(ByVal i As Long) As Long
    If i < 8 Then
        get_pulses = i
    Else
        get_pulses = (8 + (i And 7)) * celt_pow2((i \ 8) - 1)
    End If
End Function

Public Function bits2pulses(ByVal m As Long, ByVal band As Long, ByVal LM As Long, ByVal bits As Long) As Long
    Dim i As Long, lo As Long, hi As Long, cache As Long, mid As Long, clo As Long
    LM = LM + 1
    With celt_modes(m).cache
        cache = .index(LM * celt_modes(m).nbEBands + band)  ' cache = m->cache.bits + m->cache.index[...]

        lo = 0
        hi = .bits(cache)
        bits = bits - 1
        For i = 0 To LOG_MAX_PSEUDO - 1
            mid = (lo + hi + 1) \ 2
            ' OPT: Make sure this is implemented with a conditional move
            If .bits(cache + mid) >= bits Then
                hi = mid
            Else
                lo = mid
            End If
        Next
        If lo = 0 Then clo = -1 Else clo = .bits(cache + lo)
        If bits - clo <= .bits(cache + hi) - bits Then
            bits2pulses = lo
        Else
            bits2pulses = hi
        End If
    End With
End Function

Public Function pulses2bits(ByVal m As Long, ByVal band As Long, ByVal LM As Long, ByVal pulses As Long) As Long
    Dim cache As Long
    LM = LM + 1
    If pulses = 0 Then Exit Function
    cache = celt_modes(m).cache.index(LM * celt_modes(m).nbEBands + band)
    pulses2bits = celt_modes(m).cache.bits(cache + pulses) + 1
End Function

' rate.c

Private Function interp_bits2pulses(ByVal m As Long, ByVal start As Long, ByVal end_ As Long, ByVal skip_start As Long, _
        bits1() As Long, bits2() As Long, thresh() As Long, cap() As Long, ByVal total As Long, balance_ As Long, _
        ByVal skip_rsv As Long, intensity As Long, ByVal intensity_rsv As Long, dual_stereo As Long, ByVal dual_stereo_rsv As Long, bits() As Long, _
        ebits() As Long, fine_priority() As Long, ByVal C As Long, ByVal LM As Long, ec As ec_ctx, ByVal encode As Long, ByVal prev As Long, ByVal signalBandwidth As Long) As Long
    Dim psum As Long, lo As Long, hi As Long, i As Long, j As Long, logM As Long, stereo As Long, codedBands As Long
    Dim alloc_floor As Long, left_ As Long, percoeff As Long, done As Boolean, balance As Long
    Dim mid As Long, tmp As Long, band_width As Long, band_bits As Long, rem_ As Long
    Dim N0 As Long, N As Long, den As Long, offset As Long, NClogN As Long, excess As Long, bit As Long
    Dim extra_fine As Long, extra_bits As Long
    Dim eBands() As Integer

    eBands = celt_modes(m).eBands
    alloc_floor = C * 8                                 ' C<<BITRES
    If C > 1 Then stereo = 1

    logM = LM * 8                                       ' LM<<BITRES
    lo = 0
    hi = 64                                             ' 1<<ALLOC_STEPS
    For i = 0 To ALLOC_STEPS - 1
        mid = (lo + hi) \ 2
        psum = 0
        done = False
        For j = end_ - 1 To start Step -1
            tmp = bits1(j) + ((mid * bits2(j)) \ 64)    ' (mid*(opus_int32)bits2[j]>>ALLOC_STEPS)
            If tmp >= thresh(j) Or done Then
                done = True
                ' Don't allocate more than we can actually use
                If tmp < cap(j) Then psum = psum + tmp Else psum = psum + cap(j)
            Else
                If tmp >= alloc_floor Then psum = psum + alloc_floor
            End If
        Next
        If psum > total Then
            hi = mid
        Else
            lo = mid
        End If
    Next
    psum = 0
    'printf ("interp bisection gave %d\n", lo);
    done = False
    For j = end_ - 1 To start Step -1
        tmp = bits1(j) + ((lo * bits2(j)) \ 64)
        If tmp < thresh(j) And Not done Then
            If tmp >= alloc_floor Then
                tmp = alloc_floor
            Else
                tmp = 0
            End If
        Else
            done = True
        End If
        ' Don't allocate more than we can actually use
        If tmp > cap(j) Then tmp = cap(j)
        bits(j) = tmp
        psum = psum + tmp
    Next

    ' Decide which bands to skip, working backwards from the end.
    codedBands = end_
    Do
        j = codedBands - 1
        ' Never skip the first band, nor a band that has been boosted by
        '    dynalloc.
        '   In the first case, we'd be coding a bit to signal we're going to waste
        '    all the other bits.
        '   In the second case, we'd be coding a bit to redistribute all the bits
        '    we just signaled should be cocentrated in this band.
        If j <= skip_start Then
            ' Give the bit we reserved to end skipping back.
            total = total + skip_rsv
            Exit Do
        End If
        'Figure out how many left-over bits we would be adding to this band.
        '  This can include bits we've stolen back from higher, skipped bands.
        left_ = total - psum
        percoeff = left_ \ (eBands(codedBands) - eBands(start))
        left_ = left_ - (eBands(codedBands) - eBands(start)) * percoeff
        rem_ = left_ - (eBands(j) - eBands(start))
        If rem_ < 0 Then rem_ = 0
        band_width = eBands(codedBands) - eBands(j)
        band_bits = bits(j) + percoeff * band_width + rem_
        'Only code a skip decision if we're above the threshold for this band.
        '  Otherwise it is force-skipped.
        '  This ensures that we have enough bits to code the skip flag.
        tmp = thresh(j)
        If tmp < alloc_floor + 8 Then tmp = alloc_floor + 8     ' IMAX(thresh[j], alloc_floor+(1<<BITRES))
        If band_bits >= tmp Then
            If encode <> 0 Then
                ' (the encoder's skip decision: not ported)
            ElseIf ec_dec_bit_logp(ec, 1) <> 0 Then
                Exit Do
            End If
            'We used a bit to skip this band.
            psum = psum + 8
            band_bits = band_bits - 8
        End If
        'Reclaim the bits originally allocated to this band.
        psum = psum - (bits(j) + intensity_rsv)
        If intensity_rsv > 0 Then intensity_rsv = LOG2_FRAC_TABLE(j - start)
        psum = psum + intensity_rsv
        If band_bits >= alloc_floor Then
            'If we have enough for a fine energy bit per channel, use it.
            psum = psum + alloc_floor
            bits(j) = alloc_floor
        Else
            'Otherwise this band gets nothing at all.
            bits(j) = 0
        End If
        codedBands = codedBands - 1
    Loop

    ' celt_assert(codedBands > start);
    ' Code the intensity and dual stereo parameters.
    If intensity_rsv > 0 Then
        If encode <> 0 Then
            ' (the encoder's ec_enc_uint: not ported)
        Else
            intensity = start + ec_dec_uint(ec, codedBands + 1 - start)
        End If
    Else
        intensity = 0
    End If
    If intensity <= start Then
        total = total + dual_stereo_rsv
        dual_stereo_rsv = 0
    End If
    If dual_stereo_rsv > 0 Then
        If encode <> 0 Then
            ' (the encoder's ec_enc_bit_logp: not ported)
        Else
            dual_stereo = ec_dec_bit_logp(ec, 1)
        End If
    Else
        dual_stereo = 0
    End If

    ' Allocate the remaining bits
    left_ = total - psum
    percoeff = left_ \ (eBands(codedBands) - eBands(start))
    left_ = left_ - (eBands(codedBands) - eBands(start)) * percoeff
    For j = start To codedBands - 1
        bits(j) = bits(j) + percoeff * (eBands(j + 1) - eBands(j))
    Next
    For j = start To codedBands - 1
        tmp = eBands(j + 1) - eBands(j)
        If left_ < tmp Then tmp = left_                 ' IMIN(left, m->eBands[j+1]-m->eBands[j])
        bits(j) = bits(j) + tmp
        left_ = left_ - tmp
    Next
    'for (j=0;j<end;j++)printf("%d ", bits[j]);printf("\n");

    balance = 0
    For j = start To codedBands - 1
        ' celt_assert(bits[j] >= 0);
        N0 = eBands(j + 1) - eBands(j)
        N = N0 * celt_pow2(LM)
        bit = bits(j) + balance

        If N > 1 Then
            excess = bit - cap(j)
            If excess < 0 Then excess = 0
            bits(j) = bit - excess

            ' Compensate for the extra DoF in stereo
            den = C * N
            If C = 2 And N > 2 And dual_stereo = 0 And j < intensity Then den = den + 1

            NClogN = den * (celt_modes(m).logN(j) + logM)

            ' Offset for the number of fine bits by log2(N)/2 + FINE_OFFSET
            '   compared to their "fair share" of total/N
            offset = (NClogN \ 2) - den * FINE_OFFSET

            ' N=2 is the only point that doesn't match the curve
            If N = 2 Then offset = offset + (den * 8) \ 4

            ' Changing the offset for allocating the second and third
            '    fine energy bit
            If bits(j) + offset < den * 2 * 8 Then
                offset = offset + NClogN \ 4
            ElseIf bits(j) + offset < den * 3 * 8 Then
                offset = offset + NClogN \ 8
            End If

            ' Divide with rounding
            tmp = bits(j) + offset + den * 4            ' (den<<(BITRES-1))
            If tmp < 0 Then tmp = 0
            ebits(j) = (tmp \ den) \ 8                  ' celt_udiv(ebits[j], den)>>BITRES

            ' Make sure not to bust
            If C * ebits(j) > bits(j) \ 8 Then ebits(j) = (bits(j) \ celt_pow2(stereo)) \ 8

            ' More than that is useless because that's about as far as PVQ can go
            If ebits(j) > MAX_FINE_BITS Then ebits(j) = MAX_FINE_BITS

            ' If we rounded down or capped this band, make it a candidate for the
            '    final fine energy pass
            If ebits(j) * (den * 8) >= bits(j) + offset Then fine_priority(j) = 1 Else fine_priority(j) = 0

            ' Remove the allocated fine bits; the rest are assigned to PVQ
            bits(j) = bits(j) - C * ebits(j) * 8

        Else
            ' For N=1, all bits go to fine energy except for a single sign bit
            excess = bit - C * 8
            If excess < 0 Then excess = 0
            bits(j) = bit - excess
            ebits(j) = 0
            fine_priority(j) = 1
        End If

        ' Fine energy can't take advantage of the re-balancing in
        '    quant_all_bands().
        '   Instead, do the re-balancing here.
        If excess > 0 Then
            extra_fine = excess \ celt_pow2(stereo + BITRES)
            If extra_fine > MAX_FINE_BITS - ebits(j) Then extra_fine = MAX_FINE_BITS - ebits(j)
            ebits(j) = ebits(j) + extra_fine
            extra_bits = extra_fine * C * 8
            If extra_bits >= excess - balance Then fine_priority(j) = 1 Else fine_priority(j) = 0
            excess = excess - extra_bits
        End If
        balance = excess

        ' celt_assert(bits[j] >= 0);
        ' celt_assert(ebits[j] >= 0);
    Next
    ' Save any remaining bits over the cap for the rebalancing in
    '    quant_all_bands().
    balance_ = balance

    ' The skipped bands use all their bits for fine energy.
    For j = j To end_ - 1
        ebits(j) = (bits(j) \ celt_pow2(stereo)) \ 8
        ' celt_assert(C*ebits[j]<<BITRES == bits[j]);
        bits(j) = 0
        If ebits(j) < 1 Then fine_priority(j) = 1 Else fine_priority(j) = 0
    Next
    interp_bits2pulses = codedBands
End Function

Public Function clt_compute_allocation(ByVal m As Long, ByVal start As Long, ByVal end_ As Long, offsets() As Long, cap() As Long, ByVal alloc_trim As Long, intensity As Long, dual_stereo As Long, _
        ByVal total As Long, balance As Long, pulses() As Long, ebits() As Long, fine_priority() As Long, ByVal C As Long, ByVal LM As Long, ec As ec_ctx, ByVal encode As Long, ByVal prev As Long, ByVal signalBandwidth As Long) As Long
    Dim lo As Long, hi As Long, len_ As Long, j As Long, codedBands As Long, skip_start As Long
    Dim skip_rsv As Long, intensity_rsv As Long, dual_stereo_rsv As Long, done As Boolean, psum As Long, mid As Long
    Dim bitsj As Long, N As Long, bits1j As Long, bits2j As Long, pow2LM As Long
    Static bits1(0 To 24) As Long, bits2(0 To 24) As Long, thresh(0 To 24) As Long, trim_offset(0 To 24) As Long
    Dim eBands() As Integer
    If Not rate_done Then rate_init

    eBands = celt_modes(m).eBands
    pow2LM = celt_pow2(LM)
    If total < 0 Then total = 0                         ' IMAX(total, 0)
    len_ = celt_modes(m).nbEBands
    skip_start = start
    ' Reserve a bit to signal the end of manually skipped bands.
    If total >= 8 Then skip_rsv = 8 Else skip_rsv = 0
    total = total - skip_rsv
    ' Reserve bits for the intensity and dual stereo parameters.
    intensity_rsv = 0: dual_stereo_rsv = 0
    If C = 2 Then
        intensity_rsv = LOG2_FRAC_TABLE(end_ - start)
        If intensity_rsv > total Then
            intensity_rsv = 0
        Else
            total = total - intensity_rsv
            If total >= 8 Then dual_stereo_rsv = 8 Else dual_stereo_rsv = 0
            total = total - dual_stereo_rsv
        End If
    End If
    ' ALLOC(bits1, len, int); ... (VB) Static arrays of the 21 bands (and spare)

    For j = start To end_ - 1
        ' Below this threshold, we're sure not to allocate any PVQ bits
        thresh(j) = (3 * (eBands(j + 1) - eBands(j)) * pow2LM * 8) \ 16
        If thresh(j) < C * 8 Then thresh(j) = C * 8
        ' Tilt of the allocation curve
        ' (VB) >>6 of a possibly negative product: floor
        trim_offset(j) = Int(CDbl(C * (eBands(j + 1) - eBands(j)) * (alloc_trim - 5 - LM) * (end_ - j - 1) * (pow2LM * 8)) / 64)
        ' Giving less resolution to single-coefficient bands because they get
        '   more benefit from having one coarse value per coefficient
        If (eBands(j + 1) - eBands(j)) * pow2LM = 1 Then trim_offset(j) = trim_offset(j) - C * 8
    Next
    lo = 1
    hi = celt_modes(m).nbAllocVectors - 1
    Do
        done = False
        psum = 0
        mid = (lo + hi) \ 2
        For j = end_ - 1 To start Step -1
            N = eBands(j + 1) - eBands(j)
            bitsj = (C * N * celt_modes(m).allocVectors(mid * len_ + j) * pow2LM) \ 4
            If bitsj > 0 Then
                bitsj = bitsj + trim_offset(j)
                If bitsj < 0 Then bitsj = 0
            End If
            bitsj = bitsj + offsets(j)
            If bitsj >= thresh(j) Or done Then
                done = True
                ' Don't allocate more than we can actually use
                If bitsj < cap(j) Then psum = psum + bitsj Else psum = psum + cap(j)
            Else
                If bitsj >= C * 8 Then psum = psum + C * 8
            End If
        Next
        If psum > total Then
            hi = mid - 1
        Else
            lo = mid + 1
        End If
        'printf ("lo = %d, hi = %d\n", lo, hi);
    Loop While lo <= hi
    hi = lo                                             ' hi = lo--;
    lo = lo - 1
    'printf ("interp between %d and %d\n", lo, hi);
    For j = start To end_ - 1
        N = eBands(j + 1) - eBands(j)
        bits1j = (C * N * celt_modes(m).allocVectors(lo * len_ + j) * pow2LM) \ 4
        If hi >= celt_modes(m).nbAllocVectors Then
            bits2j = cap(j)
        Else
            bits2j = (C * N * celt_modes(m).allocVectors(hi * len_ + j) * pow2LM) \ 4
        End If
        If bits1j > 0 Then
            bits1j = bits1j + trim_offset(j)
            If bits1j < 0 Then bits1j = 0
        End If
        If bits2j > 0 Then
            bits2j = bits2j + trim_offset(j)
            If bits2j < 0 Then bits2j = 0
        End If
        If lo > 0 Then bits1j = bits1j + offsets(j)
        bits2j = bits2j + offsets(j)
        If offsets(j) > 0 Then skip_start = j
        bits2j = bits2j - bits1j
        If bits2j < 0 Then bits2j = 0
        bits1(j) = bits1j
        bits2(j) = bits2j
    Next
    codedBands = interp_bits2pulses(m, start, end_, skip_start, bits1, bits2, thresh, cap, _
            total, balance, skip_rsv, intensity, intensity_rsv, dual_stereo, dual_stereo_rsv, _
            pulses, ebits, fine_priority, C, LM, ec, encode, prev, signalBandwidth)
    clt_compute_allocation = codedBands
End Function
