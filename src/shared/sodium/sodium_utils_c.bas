Attribute VB_Name = "sodium_utils_c"
' (VB) What crc.cpp's crypto takes from libsodium besides the primitives (sodium_sha2_c, sodium_blake2b_c,
' sodium_chacha20_c, sodium_25519_c): randombytes_buf (randombytes/sysrandom: RtlGenRandom on Windows, as
' libsodium's), sodium_pad and sodium_unpad (sodium/utils.c, ISO/IEC 7816-4 padding), crypto_verify_16 and
' crypto_verify_32 (0 when equal), and the 32-bit and 64-bit word operations the primitives share: a
' uint32_t is a Long's bits; a uint64_t is a pair of Longs, hi and lo. Long + and - wrap (the projects have no
' overflow checks); nothing here multiplies Longs past 2^31. Not 4.7.0's code: tools/crypto_test.py checks the
' primitives against libsodium through PyNaCl.
Option Explicit

Private Declare Function RtlGenRandom Lib "advapi32" Alias "SystemFunction036" (ByRef RandomBuffer As Any, ByVal RandomBufferLength As Long) As Long

Private pow2(0 To 31) As Long           ' 1 << n
Private lowMask(0 To 32) As Long        ' (1 << n) - 1 (lowMask(32): every bit)
Private tablesReady As Boolean

Private Sub InitTables()
    Dim i As Long
    pow2(0) = 1
    For i = 1 To 30
        pow2(i) = pow2(i - 1) * 2
    Next
    pow2(31) = &H80000000
    lowMask(0) = 0
    For i = 1 To 30
        lowMask(i) = pow2(i) - 1
    Next
    lowMask(31) = &H7FFFFFFF
    lowMask(32) = -1
    tablesReady = True
End Sub

' x >> n for a uint32_t (logical), 0 <= n <= 31
Public Function sodium_shr32(ByVal x As Long, ByVal n As Long) As Long
    If Not tablesReady Then InitTables
    If n = 0 Then
        sodium_shr32 = x
    ElseIf x >= 0 Then
        sodium_shr32 = x \ pow2(n)
    Else
        sodium_shr32 = ((x And &H7FFFFFFF) \ pow2(n)) Or pow2(31 - n)
    End If
End Function

' x << n for a uint32_t (the bits past bit 31 lost), 0 <= n <= 31
Public Function sodium_shl32(ByVal x As Long, ByVal n As Long) As Long
    If Not tablesReady Then InitTables
    If n = 0 Then
        sodium_shl32 = x
    Else
        x = x And lowMask(32 - n)                           ' the bits that stay
        If (x And pow2(31 - n)) <> 0 Then
            sodium_shl32 = ((x And lowMask(31 - n)) * pow2(n)) Or &H80000000
        Else
            sodium_shl32 = x * pow2(n)
        End If
    End If
End Function

' ROTL32(x, n), 1 <= n <= 31
Public Function sodium_rotl32(ByVal x As Long, ByVal n As Long) As Long
    sodium_rotl32 = sodium_shl32(x, n) Or sodium_shr32(x, 32 - n)
End Function

' ROTR32(x, n), 1 <= n <= 31
Public Function sodium_rotr32(ByVal x As Long, ByVal n As Long) As Long
    sodium_rotr32 = sodium_shr32(x, n) Or sodium_shl32(x, 32 - n)
End Function

' (rh, rl) = (ah, al) + (bh, bl) modulo 2^64
Public Sub sodium_add64(rh As Long, rl As Long, ByVal ah As Long, ByVal al As Long, ByVal bh As Long, ByVal bl As Long)
    Dim lo As Long
    lo = al + bl
    If (lo Xor &H80000000) < (al Xor &H80000000) Then      ' unsigned lo < al: a carry out of the low word
        rh = ah + bh + 1
    Else
        rh = ah + bh
    End If
    rl = lo
End Sub

' (rh, rl) = ROTR64((h, l), n), 1 <= n <= 63
Public Sub sodium_rotr64(rh As Long, rl As Long, ByVal h As Long, ByVal l As Long, ByVal n As Long)
    Dim t As Long
    If n >= 32 Then                                         ' rotating by 32 swaps the words
        t = h
        h = l
        l = t
        n = n - 32
    End If
    If n = 0 Then
        rh = h
        rl = l
    Else
        rl = sodium_shr32(l, n) Or sodium_shl32(h, 32 - n)
        rh = sodium_shr32(h, n) Or sodium_shl32(l, 32 - n)
    End If
End Sub

' (rh, rl) = (h, l) >> n for a uint64_t, 1 <= n <= 31
Public Sub sodium_shr64(rh As Long, rl As Long, ByVal h As Long, ByVal l As Long, ByVal n As Long)
    rl = sodium_shr32(l, n) Or sodium_shl32(h, 32 - n)
    rh = sodium_shr32(h, n)
End Sub

' LOAD32_LE / STORE32_LE / LOAD32_BE / STORE32_BE
Public Function sodium_load32_le(b() As Byte, ByVal ofs As Long) As Long
    Dim t As Long
    t = b(ofs + 3)
    If t >= 128 Then
        sodium_load32_le = ((t - 128) * &H1000000) Or &H80000000 Or (CLng(b(ofs + 2)) * &H10000) Or (CLng(b(ofs + 1)) * &H100&) Or b(ofs)
    Else
        sodium_load32_le = (t * &H1000000) Or (CLng(b(ofs + 2)) * &H10000) Or (CLng(b(ofs + 1)) * &H100&) Or b(ofs)
    End If
End Function

Public Sub sodium_store32_le(b() As Byte, ByVal ofs As Long, ByVal v As Long)
    b(ofs) = v And &HFF&
    b(ofs + 1) = (v And &HFF00&) \ &H100&
    b(ofs + 2) = (v And &HFF0000) \ &H10000
    b(ofs + 3) = sodium_shr32(v, 24)
End Sub

Public Function sodium_load32_be(b() As Byte, ByVal ofs As Long) As Long
    Dim t As Long
    t = b(ofs)
    If t >= 128 Then
        sodium_load32_be = ((t - 128) * &H1000000) Or &H80000000 Or (CLng(b(ofs + 1)) * &H10000) Or (CLng(b(ofs + 2)) * &H100&) Or b(ofs + 3)
    Else
        sodium_load32_be = (t * &H1000000) Or (CLng(b(ofs + 1)) * &H10000) Or (CLng(b(ofs + 2)) * &H100&) Or b(ofs + 3)
    End If
End Function

Public Sub sodium_store32_be(b() As Byte, ByVal ofs As Long, ByVal v As Long)
    b(ofs) = sodium_shr32(v, 24)
    b(ofs + 1) = (v And &HFF0000) \ &H10000
    b(ofs + 2) = (v And &HFF00&) \ &H100&
    b(ofs + 3) = v And &HFF&
End Sub

' The uint32_t halves of a whole non-negative Double below 2^64 (a byte or bit count)
Public Sub sodium_split64(ByVal x As Double, hi As Long, lo As Long)
    Dim h As Double, l As Double
    h = Int(x / 4294967296#)
    l = x - h * 4294967296#
    If h >= 2147483648# Then hi = h - 4294967296# Else hi = h
    If l >= 2147483648# Then lo = l - 4294967296# Else lo = l
End Sub

' randombytes_buf(buf, size): RtlGenRandom, as libsodium's sysrandom on Windows (which aborts when it fails:
' sodium_misuse; here a run-time error)
Public Sub randombytes_buf(buf() As Byte, ByVal ofs As Long, ByVal size As Long)
    If size <= 0 Then Exit Sub
    If (RtlGenRandom(buf(ofs), size) And &HFF&) = 0 Then
        Err.Raise 5, "randombytes_buf", "RtlGenRandom failed"
    End If
End Sub

' crypto_verify_n: 0 when the n bytes are equal, -1 otherwise
Public Function crypto_verify_n(x() As Byte, ByVal xOfs As Long, y() As Byte, ByVal yOfs As Long, ByVal n As Long) As Long
    Dim i As Long, d As Long
    For i = 0 To n - 1
        d = d Or (x(xOfs + i) Xor y(yOfs + i))
    Next
    If d = 0 Then crypto_verify_n = 0 Else crypto_verify_n = -1
End Function

' int sodium_pad(size_t *padded_buflen_p, unsigned char *buf, size_t unpadded_buflen, size_t blocksize,
'                size_t max_buflen): buf(0 .. max_buflen - 1); the 0x80 marker at unpadded_buflen, zeros after it
'                up to the next multiple of blocksize (a whole block when already a multiple)
Public Function sodium_pad(padded_buflen As Long, buf() As Byte, ByVal unpadded_buflen As Long, ByVal blocksize As Long, ByVal max_buflen As Long) As Long
    Dim xpadlen As Long, xpadded_len As Long, i As Long
    If blocksize <= 0 Then
        sodium_pad = -1
        Exit Function
    End If
    xpadlen = blocksize - 1
    xpadlen = xpadlen - (unpadded_buflen Mod blocksize)
    xpadded_len = unpadded_buflen + xpadlen
    If xpadded_len >= max_buflen Then
        sodium_pad = -1
        Exit Function
    End If
    padded_buflen = xpadded_len + 1
    ' for i in 0 .. blocksize - 1: tail[-i] = (tail[-i] & mask) | (0x80 & barrier_mask), the mask all ones once
    ' past the barrier (i == xpadlen): the bytes after the marker zeroed, the marker written, the data kept
    For i = 0 To xpadlen - 1
        buf(xpadded_len - i) = 0
    Next
    buf(xpadded_len - xpadlen) = &H80
    sodium_pad = 0
End Function

' int sodium_unpad(size_t *unpadded_buflen_p, const unsigned char *buf, size_t padded_buflen, size_t blocksize):
' scanning back over the last block, zeros then the 0x80 marker; -1 when the last block holds no marker so placed
Public Function sodium_unpad(unpadded_buflen As Long, buf() As Byte, ByVal padded_buflen As Long, ByVal blocksize As Long) As Long
    Dim i As Long, c As Long
    If padded_buflen < blocksize Or blocksize <= 0 Then
        sodium_unpad = -1
        Exit Function
    End If
    For i = 0 To blocksize - 1
        c = buf(padded_buflen - 1 - i)
        If c <> 0 Then
            If c = &H80 Then
                unpadded_buflen = padded_buflen - 1 - i
                sodium_unpad = 0
            Else
                sodium_unpad = -1
            End If
            Exit Function
        End If
    Next
    sodium_unpad = -1
End Function
