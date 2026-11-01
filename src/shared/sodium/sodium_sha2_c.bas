Attribute VB_Name = "sodium_sha2_c"
' (VB) libsodium's crypto_hash_sha256 and crypto_hash_sha512 (crypto_hash/sha256/cp/hash_sha256_cp.c,
' crypto_hash/sha512/cp/hash_sha512_cp.c: FIPS 180-4's SHA-256 and SHA-512) with their API (init, update,
' final, and the one call): crc.cpp's sha256Sum, ed25519 (sodium_25519_c). A uint64_t word is a pair of Longs
' (sodium_utils_c); the byte count a Double (exact below 2^53). Not 4.7.0's code: tools/crypto_test.py checks
' it against libsodium.
Option Explicit

Public Type crypto_hash_sha256_state
    state(0 To 7) As Long
    count As Double                     ' bytes hashed (libsodium: uint64_t count, in bits)
    buf(0 To 63) As Byte
End Type

Public Type crypto_hash_sha512_state
    stateHi(0 To 7) As Long
    stateLo(0 To 7) As Long
    count As Double                     ' bytes hashed (libsodium: uint64_t count[2], in bits)
    buf(0 To 127) As Byte
End Type

Private K256(0 To 63) As Long
Private IV256(0 To 7) As Long
Private K512Hi(0 To 79) As Long
Private K512Lo(0 To 79) As Long
Private IV512Hi(0 To 7) As Long
Private IV512Lo(0 To 7) As Long
Private W256(0 To 63) As Long           ' (VB) the message schedules, kept between calls
Private W512Hi(0 To 79) As Long
Private W512Lo(0 To 79) As Long
Private p2(0 To 31) As Long             ' (VB) 1 << n and (1 << n) - 1, for SHA-256's own rotate and shift
Private lowMask(0 To 31) As Long
Private constantsReady As Boolean

' a uint32_t from 8 lower-case hex digits
Private Function hex32(ByVal s As String) As Long
    Dim h As Long, l As Long, i As Long
    For i = 1 To 4
        h = h * 16 + InStr(1, "0123456789abcdef", Mid$(s, i, 1), vbBinaryCompare) - 1
        l = l * 16 + InStr(1, "0123456789abcdef", Mid$(s, i + 4, 1), vbBinaryCompare) - 1
    Next
    If h >= &H8000& Then
        hex32 = ((h - &H8000&) * &H10000) Or &H80000000 Or l
    Else
        hex32 = (h * &H10000) Or l
    End If
End Function

Private Sub InitConstants()
    Dim s() As String, i As Long
    s = Split("428a2f98 71374491 b5c0fbcf e9b5dba5 3956c25b 59f111f1 923f82a4 ab1c5ed5 " & _
              "d807aa98 12835b01 243185be 550c7dc3 72be5d74 80deb1fe 9bdc06a7 c19bf174 " & _
              "e49b69c1 efbe4786 0fc19dc6 240ca1cc 2de92c6f 4a7484aa 5cb0a9dc 76f988da " & _
              "983e5152 a831c66d b00327c8 bf597fc7 c6e00bf3 d5a79147 06ca6351 14292967 " & _
              "27b70a85 2e1b2138 4d2c6dfc 53380d13 650a7354 766a0abb 81c2c92e 92722c85 " & _
              "a2bfe8a1 a81a664b c24b8b70 c76c51a3 d192e819 d6990624 f40e3585 106aa070 " & _
              "19a4c116 1e376c08 2748774c 34b0bcb5 391c0cb3 4ed8aa4a 5b9cca4f 682e6ff3 " & _
              "748f82ee 78a5636f 84c87814 8cc70208 90befffa a4506ceb bef9a3f7 c67178f2", " ")
    For i = 0 To 63
        K256(i) = hex32(s(i))
    Next
    s = Split("6a09e667 bb67ae85 3c6ef372 a54ff53a 510e527f 9b05688c 1f83d9ab 5be0cd19", " ")
    For i = 0 To 7
        IV256(i) = hex32(s(i))
    Next
    s = Split("428a2f98d728ae22 7137449123ef65cd b5c0fbcfec4d3b2f e9b5dba58189dbbc 3956c25bf348b538 " & _
              "59f111f1b605d019 923f82a4af194f9b ab1c5ed5da6d8118 d807aa98a3030242 12835b0145706fbe " & _
              "243185be4ee4b28c 550c7dc3d5ffb4e2 72be5d74f27b896f 80deb1fe3b1696b1 9bdc06a725c71235 " & _
              "c19bf174cf692694 e49b69c19ef14ad2 efbe4786384f25e3 0fc19dc68b8cd5b5 240ca1cc77ac9c65 " & _
              "2de92c6f592b0275 4a7484aa6ea6e483 5cb0a9dcbd41fbd4 76f988da831153b5 983e5152ee66dfab " & _
              "a831c66d2db43210 b00327c898fb213f bf597fc7beef0ee4 c6e00bf33da88fc2 d5a79147930aa725 " & _
              "06ca6351e003826f 142929670a0e6e70 27b70a8546d22ffc 2e1b21385c26c926 4d2c6dfc5ac42aed " & _
              "53380d139d95b3df 650a73548baf63de 766a0abb3c77b2a8 81c2c92e47edaee6 92722c851482353b " & _
              "a2bfe8a14cf10364 a81a664bbc423001 c24b8b70d0f89791 c76c51a30654be30 d192e819d6ef5218 " & _
              "d69906245565a910 f40e35855771202a 106aa07032bbd1b8 19a4c116b8d2d0c8 1e376c085141ab53 " & _
              "2748774cdf8eeb99 34b0bcb5e19b48a8 391c0cb3c5c95a63 4ed8aa4ae3418acb 5b9cca4f7763e373 " & _
              "682e6ff3d6b2b8a3 748f82ee5defb2fc 78a5636f43172f60 84c87814a1f0ab72 8cc702081a6439ec " & _
              "90befffa23631e28 a4506cebde82bde9 bef9a3f7b2c67915 c67178f2e372532b ca273eceea26619c " & _
              "d186b8c721c0c207 eada7dd6cde0eb1e f57d4f7fee6ed178 06f067aa72176fba 0a637dc5a2c898a6 " & _
              "113f9804bef90dae 1b710b35131c471b 28db77f523047d84 32caab7b40c72493 3c9ebe0a15c9bebc " & _
              "431d67c49c100d4c 4cc5d4becb3e42b6 597f299cfc657e2a 5fcb6fab3ad6faec 6c44198c4a475817", " ")
    For i = 0 To 79
        K512Hi(i) = hex32(Left$(s(i), 8))
        K512Lo(i) = hex32(Right$(s(i), 8))
    Next
    s = Split("6a09e667f3bcc908 bb67ae8584caa73b 3c6ef372fe94f82b a54ff53a5f1d36f1 " & _
              "510e527fade682d1 9b05688c2b3e6c1f 1f83d9abfb41bd6b 5be0cd19137e2179", " ")
    For i = 0 To 7
        IV512Hi(i) = hex32(Left$(s(i), 8))
        IV512Lo(i) = hex32(Right$(s(i), 8))
    Next
    p2(0) = 1
    lowMask(0) = 0
    For i = 1 To 30
        p2(i) = p2(i - 1) * 2
        lowMask(i) = p2(i) - 1
    Next
    p2(31) = &H80000000
    lowMask(31) = &H7FFFFFFF
    constantsReady = True
End Sub

' (VB) ROTR32(x, n) and x >> n (1 <= n <= 31) in one call each: SHA-256's inner loop (sodium_rotr32 is three)
Private Function rotr(ByVal x As Long, ByVal n As Long) As Long
    Dim lo As Long
    If x < 0 Then
        rotr = ((x And &H7FFFFFFF) \ p2(n)) Or p2(31 - n)
    Else
        rotr = x \ p2(n)
    End If
    lo = x And lowMask(n)                       ' the low n bits, moved to the top
    If (lo And p2(n - 1)) <> 0 Then
        rotr = rotr Or ((lo And lowMask(n - 1)) * p2(32 - n)) Or &H80000000
    Else
        rotr = rotr Or (lo * p2(32 - n))
    End If
End Function

Private Function shr(ByVal x As Long, ByVal n As Long) As Long
    If x < 0 Then
        shr = ((x And &H7FFFFFFF) \ p2(n)) Or p2(31 - n)
    Else
        shr = x \ p2(n)
    End If
End Function

' SHA-512's initial hash value, word i (BLAKE2b's IV is the same)
Public Sub sodium_sha512_iv(ByVal i As Long, hi As Long, lo As Long)
    If Not constantsReady Then InitConstants
    hi = IV512Hi(i)
    lo = IV512Lo(i)
End Sub

'--------------------------------------------------------------------------------------------------------
' SHA-256

Private Sub sha256_transform(st As crypto_hash_sha256_state, src() As Byte, ByVal ofs As Long)
    Dim t As Long, a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, h As Long
    Dim s0 As Long, s1 As Long, t1 As Long, t2 As Long, x As Long
    For t = 0 To 15
        W256(t) = sodium_load32_be(src, ofs + 4 * t)
    Next
    For t = 16 To 63
        x = W256(t - 15)
        s0 = rotr(x, 7) Xor rotr(x, 18) Xor shr(x, 3)
        x = W256(t - 2)
        s1 = rotr(x, 17) Xor rotr(x, 19) Xor shr(x, 10)
        W256(t) = W256(t - 16) + s0 + W256(t - 7) + s1
    Next
    a = st.state(0): b = st.state(1): c = st.state(2): d = st.state(3)
    e = st.state(4): f = st.state(5): g = st.state(6): h = st.state(7)
    For t = 0 To 63
        s1 = rotr(e, 6) Xor rotr(e, 11) Xor rotr(e, 25)
        t1 = h + s1 + ((e And f) Xor ((Not e) And g)) + K256(t) + W256(t)
        s0 = rotr(a, 2) Xor rotr(a, 13) Xor rotr(a, 22)
        t2 = s0 + ((a And b) Xor (a And c) Xor (b And c))
        h = g
        g = f
        f = e
        e = d + t1
        d = c
        c = b
        b = a
        a = t1 + t2
    Next
    st.state(0) = st.state(0) + a: st.state(1) = st.state(1) + b
    st.state(2) = st.state(2) + c: st.state(3) = st.state(3) + d
    st.state(4) = st.state(4) + e: st.state(5) = st.state(5) + f
    st.state(6) = st.state(6) + g: st.state(7) = st.state(7) + h
End Sub

Public Function crypto_hash_sha256_init(st As crypto_hash_sha256_state) As Long
    Dim i As Long
    If Not constantsReady Then InitConstants
    For i = 0 To 7
        st.state(i) = IV256(i)
    Next
    st.count = 0
End Function

Public Function crypto_hash_sha256_update(st As crypto_hash_sha256_state, inBuf() As Byte, ByVal ofs As Long, ByVal inlen As Long) As Long
    Dim r As Long, i As Long, n As Long
    If inlen <= 0 Then Exit Function
    r = st.count - Int(st.count / 64#) * 64#
    st.count = st.count + inlen
    If r > 0 Then
        n = 64 - r
        If n > inlen Then n = inlen
        For i = 0 To n - 1
            st.buf(r + i) = inBuf(ofs + i)
        Next
        ofs = ofs + n
        inlen = inlen - n
        If r + n < 64 Then Exit Function
        sha256_transform st, st.buf, 0
    End If
    Do While inlen >= 64
        sha256_transform st, inBuf, ofs
        ofs = ofs + 64
        inlen = inlen - 64
    Loop
    For i = 0 To inlen - 1
        st.buf(i) = inBuf(ofs + i)
    Next
End Function

Public Function crypto_hash_sha256_final(st As crypto_hash_sha256_state, out() As Byte, ByVal outOfs As Long) As Long
    Dim r As Long, i As Long, bitsHi As Long, bitsLo As Long
    r = st.count - Int(st.count / 64#) * 64#
    sodium_split64 st.count * 8#, bitsHi, bitsLo
    st.buf(r) = &H80
    r = r + 1
    If r > 56 Then
        For i = r To 63
            st.buf(i) = 0
        Next
        sha256_transform st, st.buf, 0
        r = 0
    End If
    For i = r To 55
        st.buf(i) = 0
    Next
    sodium_store32_be st.buf, 56, bitsHi
    sodium_store32_be st.buf, 60, bitsLo
    sha256_transform st, st.buf, 0
    For i = 0 To 7
        sodium_store32_be out, outOfs + 4 * i, st.state(i)
    Next
End Function

' crypto_hash_sha256(out, in, inlen): out(outOfs .. outOfs + 31)
Public Function crypto_hash_sha256(out() As Byte, ByVal outOfs As Long, inBuf() As Byte, ByVal ofs As Long, ByVal inlen As Long) As Long
    Dim st As crypto_hash_sha256_state
    crypto_hash_sha256_init st
    crypto_hash_sha256_update st, inBuf, ofs, inlen
    crypto_hash_sha256_final st, out, outOfs
End Function

'--------------------------------------------------------------------------------------------------------
' SHA-512

Private Sub sha512_transform(st As crypto_hash_sha512_state, src() As Byte, ByVal ofs As Long)
    Dim t As Long, i As Long
    Dim ah As Long, al As Long, bh As Long, bl As Long, ch As Long, cl As Long, dh As Long, dl As Long
    Dim eh As Long, el As Long, fh As Long, fl As Long, gh As Long, gl As Long, hh As Long, hl As Long
    Dim xh As Long, xl As Long, yh As Long, yl As Long, zh As Long, zl As Long
    Dim t1h As Long, t1l As Long, t2h As Long, t2l As Long
    For t = 0 To 15
        W512Hi(t) = sodium_load32_be(src, ofs + 8 * t)
        W512Lo(t) = sodium_load32_be(src, ofs + 8 * t + 4)
    Next
    For t = 16 To 79
        ' s0 = ROTR(w[t-15], 1) ^ ROTR(w[t-15], 8) ^ SHR(w[t-15], 7)
        sodium_rotr64 xh, xl, W512Hi(t - 15), W512Lo(t - 15), 1
        sodium_rotr64 yh, yl, W512Hi(t - 15), W512Lo(t - 15), 8
        sodium_shr64 zh, zl, W512Hi(t - 15), W512Lo(t - 15), 7
        t1h = xh Xor yh Xor zh
        t1l = xl Xor yl Xor zl
        ' s1 = ROTR(w[t-2], 19) ^ ROTR(w[t-2], 61) ^ SHR(w[t-2], 6)
        sodium_rotr64 xh, xl, W512Hi(t - 2), W512Lo(t - 2), 19
        sodium_rotr64 yh, yl, W512Hi(t - 2), W512Lo(t - 2), 61
        sodium_shr64 zh, zl, W512Hi(t - 2), W512Lo(t - 2), 6
        t2h = xh Xor yh Xor zh
        t2l = xl Xor yl Xor zl
        sodium_add64 xh, xl, W512Hi(t - 16), W512Lo(t - 16), t1h, t1l
        sodium_add64 xh, xl, xh, xl, W512Hi(t - 7), W512Lo(t - 7)
        sodium_add64 W512Hi(t), W512Lo(t), xh, xl, t2h, t2l
    Next
    ah = st.stateHi(0): al = st.stateLo(0): bh = st.stateHi(1): bl = st.stateLo(1)
    ch = st.stateHi(2): cl = st.stateLo(2): dh = st.stateHi(3): dl = st.stateLo(3)
    eh = st.stateHi(4): el = st.stateLo(4): fh = st.stateHi(5): fl = st.stateLo(5)
    gh = st.stateHi(6): gl = st.stateLo(6): hh = st.stateHi(7): hl = st.stateLo(7)
    For t = 0 To 79
        ' t1 = h + S1(e) + Ch(e, f, g) + K[t] + W[t]; S1(e) = ROTR(e, 14) ^ ROTR(e, 18) ^ ROTR(e, 41)
        sodium_rotr64 xh, xl, eh, el, 14
        sodium_rotr64 yh, yl, eh, el, 18
        sodium_rotr64 zh, zl, eh, el, 41
        sodium_add64 t1h, t1l, hh, hl, xh Xor yh Xor zh, xl Xor yl Xor zl
        sodium_add64 t1h, t1l, t1h, t1l, (eh And fh) Xor ((Not eh) And gh), (el And fl) Xor ((Not el) And gl)
        sodium_add64 t1h, t1l, t1h, t1l, K512Hi(t), K512Lo(t)
        sodium_add64 t1h, t1l, t1h, t1l, W512Hi(t), W512Lo(t)
        ' t2 = S0(a) + Maj(a, b, c); S0(a) = ROTR(a, 28) ^ ROTR(a, 34) ^ ROTR(a, 39)
        sodium_rotr64 xh, xl, ah, al, 28
        sodium_rotr64 yh, yl, ah, al, 34
        sodium_rotr64 zh, zl, ah, al, 39
        sodium_add64 t2h, t2l, xh Xor yh Xor zh, xl Xor yl Xor zl, _
                     (ah And bh) Xor (ah And ch) Xor (bh And ch), (al And bl) Xor (al And cl) Xor (bl And cl)
        hh = gh: hl = gl
        gh = fh: gl = fl
        fh = eh: fl = el
        sodium_add64 eh, el, dh, dl, t1h, t1l
        dh = ch: dl = cl
        ch = bh: cl = bl
        bh = ah: bl = al
        sodium_add64 ah, al, t1h, t1l, t2h, t2l
    Next
    sodium_add64 st.stateHi(0), st.stateLo(0), st.stateHi(0), st.stateLo(0), ah, al
    sodium_add64 st.stateHi(1), st.stateLo(1), st.stateHi(1), st.stateLo(1), bh, bl
    sodium_add64 st.stateHi(2), st.stateLo(2), st.stateHi(2), st.stateLo(2), ch, cl
    sodium_add64 st.stateHi(3), st.stateLo(3), st.stateHi(3), st.stateLo(3), dh, dl
    sodium_add64 st.stateHi(4), st.stateLo(4), st.stateHi(4), st.stateLo(4), eh, el
    sodium_add64 st.stateHi(5), st.stateLo(5), st.stateHi(5), st.stateLo(5), fh, fl
    sodium_add64 st.stateHi(6), st.stateLo(6), st.stateHi(6), st.stateLo(6), gh, gl
    sodium_add64 st.stateHi(7), st.stateLo(7), st.stateHi(7), st.stateLo(7), hh, hl
End Sub

Public Function crypto_hash_sha512_init(st As crypto_hash_sha512_state) As Long
    Dim i As Long
    If Not constantsReady Then InitConstants
    For i = 0 To 7
        st.stateHi(i) = IV512Hi(i)
        st.stateLo(i) = IV512Lo(i)
    Next
    st.count = 0
End Function

Public Function crypto_hash_sha512_update(st As crypto_hash_sha512_state, inBuf() As Byte, ByVal ofs As Long, ByVal inlen As Long) As Long
    Dim r As Long, i As Long, n As Long
    If inlen <= 0 Then Exit Function
    r = st.count - Int(st.count / 128#) * 128#
    st.count = st.count + inlen
    If r > 0 Then
        n = 128 - r
        If n > inlen Then n = inlen
        For i = 0 To n - 1
            st.buf(r + i) = inBuf(ofs + i)
        Next
        ofs = ofs + n
        inlen = inlen - n
        If r + n < 128 Then Exit Function
        sha512_transform st, st.buf, 0
    End If
    Do While inlen >= 128
        sha512_transform st, inBuf, ofs
        ofs = ofs + 128
        inlen = inlen - 128
    Loop
    For i = 0 To inlen - 1
        st.buf(i) = inBuf(ofs + i)
    Next
End Function

Public Function crypto_hash_sha512_final(st As crypto_hash_sha512_state, out() As Byte, ByVal outOfs As Long) As Long
    Dim r As Long, i As Long, bitsHi As Long, bitsLo As Long
    r = st.count - Int(st.count / 128#) * 128#
    sodium_split64 st.count * 8#, bitsHi, bitsLo
    st.buf(r) = &H80
    r = r + 1
    If r > 112 Then
        For i = r To 127
            st.buf(i) = 0
        Next
        sha512_transform st, st.buf, 0
        r = 0
    End If
    For i = r To 119                        ' the 128-bit length's high 64 bits are 0 (below 2^61 bytes)
        st.buf(i) = 0
    Next
    sodium_store32_be st.buf, 120, bitsHi
    sodium_store32_be st.buf, 124, bitsLo
    sha512_transform st, st.buf, 0
    For i = 0 To 7
        sodium_store32_be out, outOfs + 8 * i, st.stateHi(i)
        sodium_store32_be out, outOfs + 8 * i + 4, st.stateLo(i)
    Next
End Function

' crypto_hash_sha512(out, in, inlen): out(outOfs .. outOfs + 63)
Public Function crypto_hash_sha512(out() As Byte, ByVal outOfs As Long, inBuf() As Byte, ByVal ofs As Long, ByVal inlen As Long) As Long
    Dim st As crypto_hash_sha512_state
    crypto_hash_sha512_init st
    crypto_hash_sha512_update st, inBuf, ofs, inlen
    crypto_hash_sha512_final st, out, outOfs
End Function
