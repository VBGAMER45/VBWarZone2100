Attribute VB_Name = "entcode_c"
' ref/opus/celt (libopus 1.5.2): entcode.h, entcode.c, entdec.c, entdec.h, with mfrngcod.h's constants and
' ecintrin.h's EC_ILOG and EC_MINI: the range decoder (cluster K, fork O2; O3's SILK and opus_decoder.c call
' it too: K_DESIGN.md's contract). The encoder (entenc.c, entenc.h) is not ported: nothing in the decoder
' reaches it. USE_SMALL_DIV_TABLE is undefined (it is OPUS_ARM_ASM's), so SMALL_DIV_TABLE and celt_udiv's table
' path are left out; celt_udiv and celt_sudiv are the plain divisions (written in place by their callers).
' (VB) The struct is ec_ctx with buf() a copy of the storage bytes the C points at (ec_dec_init copies them);
' an opus_uint32 field (rng, val, ext, end_window, ...) is a Long holding its bits. In the decoder val is
' below 2^31 and rng in (2^23, 2^31]: rng = 2^31 is the one negative Long, and the code below handles it where
' a shift or subtraction would need the unsigned value. EC_CLZ: ec_ilog is the C's fallback (same results).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

'OPT: ec_window must be at least 32 bits, but if you have fast arithmetic on a
'   larger type, you can speed up the decoder by using it here.
' typedef opus_uint32 ec_window: (VB) a Long's bits

Public Const EC_WINDOW_SIZE As Long = 32

'The number of bits to use for the range-coded part of unsigned integers.
Public Const EC_UINT_BITS As Long = 8

'The resolution of fractional-precision bit usage measurements, i.e.,
'   3 => 1/8th bits.
Public Const BITRES As Long = 3

' mfrngcod.h: Constants used by the entropy encoder/decoder.
'The number of bits to output at a time.
Private Const EC_SYM_BITS As Long = 8
'The total number of bits in each of the state registers.
Private Const EC_CODE_BITS As Long = 32
'The maximum symbol value.
Private Const EC_SYM_MAX As Long = 255
'Bits to shift by to move a symbol into the high-order position.
Private Const EC_CODE_SHIFT As Long = 23            ' EC_CODE_BITS-EC_SYM_BITS-1
'Carry bit of the high-order range symbol.
Private Const EC_CODE_TOP As Long = &H80000000      ' 1U<<(EC_CODE_BITS-1): (VB) the bits of 2^31
'Low-order bit of the high-order range symbol.
Private Const EC_CODE_BOT As Long = &H800000        ' EC_CODE_TOP>>EC_SYM_BITS
'The number of bits available for the last, partial symbol in the code field.
Private Const EC_CODE_EXTRA As Long = 7             ' (EC_CODE_BITS-2)%EC_SYM_BITS+1

'The entropy encoder/decoder context.
'  We use the same structure for both, so that common functions like ec_tell()
'   can be used on either one.
Public Type ec_ctx
    'Buffered input/output.
    buf() As Byte                   ' unsigned char *buf: (VB) a copy of the storage bytes the C points at
    'The size of the buffer.
    storage As Long                 ' opus_uint32
    'The offset at which the last byte containing raw bits was read/written.
    end_offs As Long                ' opus_uint32
    'Bits that will be read from/written at the end.
    end_window As Long              ' ec_window
    'Number of valid bits in end_window.
    nend_bits As Long
    'The total number of whole bits read/written.
    '  This does not include partial bits currently in the range coder.
    nbits_total As Long
    'The offset at which the next range coder byte will be read/written.
    offs As Long                    ' opus_uint32
    'The number of values in the current range.
    rng As Long                     ' opus_uint32
    'In the decoder: the difference between the top of the current range and
    '   the input value, minus one.
    '  In the encoder: the low end of the current range.
    val As Long                     ' opus_uint32
    'In the decoder: the saved normalization factor from ec_decode().
    '  In the encoder: the number of oustanding carry propagating symbols.
    ext As Long                     ' opus_uint32
    'A buffered input/output symbol, awaiting carry propagation.
    rem_ As Long                    ' rem
    'Nonzero if an error occurred.
    error_ As Long                  ' error
End Type

' (VB) 2^n and 2^n - 1 as Longs (ec_pow2(31) is the bits of 2^31), ec_tell_frac's correction table
Private ec_pow2(0 To 31) As Long
Private ec_mask(0 To 31) As Long
Private ec_correction(0 To 7) As Long
Private ec_tables_done As Boolean

Private Sub ec_tables_init()
    Dim i As Long
    ec_pow2(0) = 1
    For i = 1 To 30
        ec_pow2(i) = ec_pow2(i - 1) * 2
    Next
    ec_pow2(31) = &H80000000
    For i = 0 To 30
        ec_mask(i) = ec_pow2(i) - 1
    Next
    ec_mask(31) = &H7FFFFFFF
    ' static const unsigned correction[8] = {35733, 38967, 42495, 46340, 50535, 55109, 60097, 65535};
    ec_correction(0) = 35733: ec_correction(1) = 38967: ec_correction(2) = 42495: ec_correction(3) = 46340
    ec_correction(4) = 50535: ec_correction(5) = 55109: ec_correction(6) = 60097: ec_correction(7) = 65535
    ec_tables_done = True
End Sub

' (VB) C's x >> n for an opus_uint32 x (a Long's bits), 0 <= n <= 31
Private Function ec_ushr(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then
        ec_ushr = x
    ElseIf x >= 0 Then
        ec_ushr = x \ ec_pow2(n)
    Else
        ec_ushr = ((x And &H7FFFFFFF) \ ec_pow2(n)) Or ec_pow2(31 - n)
    End If
End Function

' entcode.h

Public Function ec_range_bytes(dec As ec_ctx) As Long
    ec_range_bytes = dec.offs
End Function

' ec_get_buffer: (VB) the C returns the buffer pointer; callers use dec.buf

Public Function ec_get_error(dec As ec_ctx) As Long
    ec_get_error = dec.error_
End Function

'Returns the number of bits "used" by the encoded or decoded symbols so far.
'  This same number can be computed in either the encoder or the decoder, and is
'   suitable for making coding decisions.
'  Return: The number of bits.
'          This will always be slightly larger than the exact value (e.g., all
'           rounding error is in the positive direction).
Public Function ec_tell(dec As ec_ctx) As Long
    ec_tell = dec.nbits_total - ec_ilog(dec.rng)
End Function

' entcode.c

' #if !defined(EC_CLZ)
'This is a fallback for systems where we don't know how to access
'   a BSR or CLZ instruction (see ecintrin.h).
'  If you are optimizing Opus on a new platform and it has a native CLZ or
'   BZR (e.g. cell, MIPS, x86, etc) then making it available to Opus will be
'   an easy performance win.
' (VB) EC_ILOG(_x): the number of bits of the uint32 _v (0 for 0); the C's branchless version gives the same
Public Function ec_ilog(ByVal v As Long) As Long
    'On a Pentium M, this branchless version tested as the fastest on
    '   1,000,000,000 random 32-bit integers, edging out a similar version with
    '   branches, and a 256-entry LUT version.
    Dim ret As Long
    If v < 0 Then ec_ilog = 32: Exit Function
    If v = 0 Then Exit Function
    ret = 1
    If v >= &H10000 Then ret = 17: v = v \ &H10000
    If v >= &H100 Then ret = ret + 8: v = v \ &H100
    If v >= &H10 Then ret = ret + 4: v = v \ &H10
    If v >= 4 Then ret = ret + 2: v = v \ 4
    If v >= 2 Then ret = ret + 1
    ec_ilog = ret
End Function

' #if 1
' This is a faster version of ec_tell_frac() that takes advantage
'   of the low (1/8 bit) resolution to use just a linear function
'   followed by a lookup to determine the exact transition thresholds.
Public Function ec_tell_frac(dec As ec_ctx) As Long
    Dim nbits As Long, r As Long, l As Long, b As Long
    If Not ec_tables_done Then ec_tables_init
    nbits = dec.nbits_total * 8                 ' <<BITRES
    l = ec_ilog(dec.rng)
    r = ec_ushr(dec.rng, l - 16)
    b = (r \ 4096) - 8
    If r > ec_correction(b) Then b = b + 1
    l = l * 8 + b
    ec_tell_frac = nbits - l
End Function
' #else: the exact version (not compiled)

' entdec.c

'A range decoder.
'  This is an entropy decoder based upon \cite{Mar79}, which is itself a
'   rediscovery of the FIFO arithmetic code introduced by \cite{Pas76}.
'  It is very similar to arithmetic encoding, except that encoding is done with
'   digits in any base, instead of with bits, and so it is faster when using
'   larger bases (i.e.: a byte).
'  The author claims an average waste of $\frac{1}{2}\log_b(2b)$ bits, where $b$
'   is the base, longer than the theoretical optimum, but to my knowledge there
'   is no published justification for this claim.
'  This only seems true when using near-infinite precision arithmetic so that
'   the process is carried out with no rounding errors.
'
'  An excellent description of implementation details is available at
'   http://www.arturocampos.com/ac_range.html
'  A recent work \cite{MNW98} which proposes several changes to arithmetic
'   encoding for efficiency actually re-discovers many of the principles
'   behind range encoding, and presents a good theoretical analysis of them.
'
'  End of stream is handled by writing out the smallest number of bits that
'   ensures that the stream will be correctly decoded regardless of the value of
'   any subsequent bits.
'  ec_tell() can be used to determine how many bits were needed to decode
'   all the symbols thus far; other data can be packed in the remaining bits of
'   the input buffer.
'  (the references Pas76, Mar79, MNW98: see entdec.c)

Private Function ec_read_byte(dec As ec_ctx) As Long
    If dec.offs < dec.storage Then
        ec_read_byte = dec.buf(dec.offs)
        dec.offs = dec.offs + 1
    End If
End Function

Private Function ec_read_byte_from_end(dec As ec_ctx) As Long
    If dec.end_offs < dec.storage Then
        dec.end_offs = dec.end_offs + 1
        ec_read_byte_from_end = dec.buf(dec.storage - dec.end_offs)
    End If
End Function

'Normalizes the contents of val and rng so that rng lies entirely in the
'   high-order symbol.
Private Sub ec_dec_normalize(dec As ec_ctx)
    Dim sym As Long
    'If the range is too small, rescale it and input some bits.
    Do While dec.rng <= EC_CODE_BOT And dec.rng > 0     ' (VB) rng = 2^31 is a negative Long: not too small
        dec.nbits_total = dec.nbits_total + EC_SYM_BITS
        ' rng<<=EC_SYM_BITS
        If dec.rng = EC_CODE_BOT Then dec.rng = EC_CODE_TOP Else dec.rng = dec.rng * 256
        'Use up the remaining bits from our last symbol.
        sym = dec.rem_
        'Read the next value from the input.
        dec.rem_ = ec_read_byte(dec)
        'Take the rest of the bits we need from this new symbol.
        sym = (sym * 256 Or dec.rem_) \ 2               ' >>(EC_SYM_BITS-EC_CODE_EXTRA)
        'And subtract them from val, capped to be less than EC_CODE_TOP.
        dec.val = ((dec.val And &H7FFFFF) * 256 + (EC_SYM_MAX And Not sym)) And &H7FFFFFFF
    Loop
End Sub

' (VB) unsigned char *_buf is buf() from bufOffset: its storage bytes are copied into dec.buf
Public Sub ec_dec_init(dec As ec_ctx, buf() As Byte, ByVal bufOffset As Long, ByVal storage As Long)
    If Not ec_tables_done Then ec_tables_init
    If storage > 0 Then
        ReDim dec.buf(0 To storage - 1)
        CopyMemory dec.buf(0), buf(bufOffset), storage
    Else
        ReDim dec.buf(0 To 0)
    End If
    dec.storage = storage
    dec.end_offs = 0
    dec.end_window = 0
    dec.nend_bits = 0
    'This is the offset from which ec_tell() will subtract partial bits.
    '  The final value after the ec_dec_normalize() call will be the same as in
    '   the encoder, but we have to compensate for the bits that are added there.
    dec.nbits_total = EC_CODE_BITS + 1 - ((EC_CODE_BITS - EC_CODE_EXTRA) \ EC_SYM_BITS) * EC_SYM_BITS
    dec.offs = 0
    dec.rng = 128                                       ' 1U<<EC_CODE_EXTRA
    dec.rem_ = ec_read_byte(dec)
    dec.val = dec.rng - 1 - (dec.rem_ \ 2)              ' _this->rem>>(EC_SYM_BITS-EC_CODE_EXTRA)
    dec.error_ = 0
    'Normalize the interval.
    ec_dec_normalize dec
End Sub

Public Function ec_decode(dec As ec_ctx, ByVal ft As Long) As Long
    Dim s As Long
    ' celt_udiv(_this->rng,_ft)
    If dec.rng < 0 Then dec.ext = Int(2147483648# / ft) Else dec.ext = dec.rng \ ft
    s = dec.val \ dec.ext
    ' _ft-EC_MINI(s+1,_ft)
    If s + 1 < ft Then ec_decode = ft - (s + 1)
End Function

Public Function ec_decode_bin(dec As ec_ctx, ByVal bits As Long) As Long
    Dim s As Long
    dec.ext = ec_ushr(dec.rng, bits)
    s = dec.val \ dec.ext
    ' (1U<<_bits)-EC_MINI(s+1U,1U<<_bits)
    If s + 1 < ec_pow2(bits) Then ec_decode_bin = ec_pow2(bits) - (s + 1)
End Function

Public Sub ec_dec_update(dec As ec_ctx, ByVal fl As Long, ByVal fh As Long, ByVal ft As Long)
    Dim s As Long
    s = dec.ext * (ft - fh)                             ' IMUL32: below rng, so no overflow
    dec.val = dec.val - s
    If fl > 0 Then
        dec.rng = dec.ext * (fh - fl)
    ElseIf dec.rng < 0 Then
        If s <> 0 Then dec.rng = (&H7FFFFFFF - s) + 1   ' (VB) 2^31 - s
    Else
        dec.rng = dec.rng - s
    End If
    ec_dec_normalize dec
End Sub

'The probability of having a "one" is 1/(1<<_logp).
Public Function ec_dec_bit_logp(dec As ec_ctx, ByVal logp As Long) As Long
    Dim r As Long, d As Long, s As Long, ret As Long
    r = dec.rng
    d = dec.val
    s = ec_ushr(r, logp)
    If d < s Then ret = 1
    If ret = 0 Then dec.val = d - s
    If ret <> 0 Then
        dec.rng = s
    ElseIf r < 0 Then
        dec.rng = (&H7FFFFFFF - s) + 1                  ' (VB) 2^31 - s (s >= 1)
    Else
        dec.rng = r - s
    End If
    ec_dec_normalize dec
    ec_dec_bit_logp = ret
End Function

' (VB) const unsigned char *_icdf is icdf() from icdfOffset
Public Function ec_dec_icdf(dec As ec_ctx, icdf() As Byte, ByVal icdfOffset As Long, ByVal ftb As Long) As Long
    Dim r As Long, d As Long, s As Long, t As Long, ret As Long
    s = dec.rng
    d = dec.val
    r = ec_ushr(s, ftb)
    ret = -1
    Do
        t = s
        ret = ret + 1
        s = r * icdf(icdfOffset + ret)
    Loop While d < s
    dec.val = d - s
    If t >= 0 Then
        dec.rng = t - s
    ElseIf s <> 0 Then
        dec.rng = (&H7FFFFFFF - s) + 1                  ' (VB) 2^31 - s
    Else
        dec.rng = t
    End If
    ec_dec_normalize dec
    ec_dec_icdf = ret
End Function

' (VB) const opus_uint16 *_icdf is icdf() (Longs) from icdfOffset
Public Function ec_dec_icdf16(dec As ec_ctx, icdf() As Long, ByVal icdfOffset As Long, ByVal ftb As Long) As Long
    Dim r As Long, d As Long, s As Long, t As Long, ret As Long
    s = dec.rng
    d = dec.val
    r = ec_ushr(s, ftb)
    ret = -1
    Do
        t = s
        ret = ret + 1
        s = r * icdf(icdfOffset + ret)
    Loop While d < s
    dec.val = d - s
    If t >= 0 Then
        dec.rng = t - s
    ElseIf s <> 0 Then
        dec.rng = (&H7FFFFFFF - s) + 1                  ' (VB) 2^31 - s
    Else
        dec.rng = t
    End If
    ec_dec_normalize dec
    ec_dec_icdf16 = ret
End Function

' (VB) _ft and the result are opus_uint32 (a Long's bits)
Public Function ec_dec_uint(dec As ec_ctx, ByVal ft As Long) As Long
    Dim ft_ As Long, s As Long, ftb As Long, t As Double, ftd As Double
    'In order to optimize EC_ILOG(), it is undefined for the value 0.
    ' celt_assert(_ft>1);
    If ft = &H80000000 Then ft = &H7FFFFFFF Else ft = ft - 1        ' _ft--
    ftb = ec_ilog(ft)
    If ftb > EC_UINT_BITS Then
        ftb = ftb - EC_UINT_BITS
        ft_ = ec_ushr(ft, ftb) + 1
        s = ec_decode(dec, ft_)
        ec_dec_update dec, s, s + 1, ft_
        ' t=(opus_uint32)s<<ftb|ec_dec_bits(_this,ftb): (VB) in a Double (below 2^32)
        t = CDbl(s) * ec_pow2(ftb) + ec_dec_bits(dec, ftb)
        ftd = ft
        If ft < 0 Then ftd = ftd + 4294967296#
        If t <= ftd Then
            If t >= 2147483648# Then t = t - 4294967296#
            ec_dec_uint = t
            Exit Function
        End If
        dec.error_ = 1
        ec_dec_uint = ft
    Else
        ft = ft + 1                                     ' _ft++
        s = ec_decode(dec, ft)
        ec_dec_update dec, s, s + 1, ft
        ec_dec_uint = s
    End If
End Function

' (VB) the result is an opus_uint32 (a Long's bits; _bits <= 31)
Public Function ec_dec_bits(dec As ec_ctx, ByVal bits As Long) As Long
    Dim window As Long, available As Long, ret As Long, b As Long
    window = dec.end_window
    available = dec.nend_bits
    If available < bits Then
        Do
            ' window|=(ec_window)ec_read_byte_from_end(_this)<<available;
            b = ec_read_byte_from_end(dec)
            If available = 24 Then
                If b >= 128 Then window = window Or ((b And 127) * &H1000000) Or &H80000000 Else window = window Or (b * &H1000000)
            Else
                window = window Or (b * ec_pow2(available))
            End If
            available = available + EC_SYM_BITS
        Loop While available <= EC_WINDOW_SIZE - EC_SYM_BITS
    End If
    ret = window And ec_mask(bits)
    window = ec_ushr(window, bits)
    available = available - bits
    dec.end_window = window
    dec.nend_bits = available
    dec.nbits_total = dec.nbits_total + bits
    ec_dec_bits = ret
End Function
