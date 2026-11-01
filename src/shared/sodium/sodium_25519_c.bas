Attribute VB_Name = "sodium_25519_c"
' (VB) libsodium's Curve25519 and Ed25519 functions that crc.cpp calls: crypto_scalarmult_curve25519 (X25519),
' crypto_sign_ed25519_seed_keypair / _keypair / _detached / _verify_detached / _sk_to_pk /
' _pk_to_curve25519 / _sk_to_curve25519, crypto_core_ed25519_is_valid_point, crypto_kx_client_session_keys
' and crypto_kx_server_session_keys. The arithmetic is TweetNaCl's (public domain): a field element is 16
' limbs of 16 bits (libsodium's ref10 uses 25.5-bit limbs and 64-bit products, which VB has no integer for),
' each limb a Double: every value stays an integer below 2^53, so Double arithmetic is exact (worked out
' where M is: limbs in at most 2^17.6, so a product below 2^35.2, 16 of them below 2^39.2, times 39 below
' 2^44.5; the scalar reduction modL below 2^34); a floor shift is Int(x / 2^n), exact for a power of 2.
' libsodium's acceptance rules are kept as libsodium 1.0.20 has them (vcpkg's, which 4.7.0 builds with): the
' canonical encodings, the small-order and main-subgroup checks, X25519's small-order blacklist and its
' all-zero result. Swaps are plain (TweetNaCl's constant-time masks protect against timing, which a port
' that is checked bit for bit against libsodium does not need). Not 4.7.0's code: tools/crypto_test.py checks
' every function against libsodium through PyNaCl.
Option Explicit

Private Type gf
    v(0 To 15) As Double
End Type

Private Type ge                         ' TweetNaCl's gf p[4]: extended coordinates X, Y, Z, T
    X As gf
    Y As gf
    Z As gf
    T As gf
End Type

Private gf0 As gf
Private gf1 As gf
Private c121665 As gf                   ' {0xDB41, 1}
Private EDWARDS_D As gf                         ' -121665 / 121666
Private EDWARDS_D2 As gf                        ' 2 * EDWARDS_D
Private SQRTM1 As gf                    ' TweetNaCl's I: 2^((p - 1) / 4), a square root of -1
Private BASEX As gf                     ' the base point (TweetNaCl's X, Y)
Private BASEY As gf
Private Lc(0 To 31) As Double           ' the group order L, little-endian bytes
Private Lbytes(0 To 31) As Byte
Private mt(0 To 30) As Double           ' (VB) M's product, kept between calls
Private bitv(0 To 7) As Long            ' 1 << n for a byte's bits
Private ready As Boolean

Private Const crypto_scalarmult_curve25519_BYTES As Long = 32
Private Const crypto_kx_SESSIONKEYBYTES As Long = 32

#If SODIUM_LIMB_CHECK Then
' (VB) CryptoTests' check of the exactness argument above (its project sets SODIUM_LIMB_CHECK): the largest
' magnitude any limb, product sum or scalar digit reaches
Private limbMax As Double

Private Sub limbSee(ByVal v As Double)
    If Abs(v) > limbMax Then limbMax = Abs(v)
End Sub

Public Function sodium_25519_limbMax() As Double
    sodium_25519_limbMax = limbMax
End Function
#End If

'--------------------------------------------------------------------------------------------------------
' the field: integers modulo p = 2^255 - 19

Private Sub car25519(o As gf)
    Dim i As Long, c As Double
    For i = 0 To 15
#If SODIUM_LIMB_CHECK Then
        limbSee o.v(i)
#End If
        o.v(i) = o.v(i) + 65536#
        c = Int(o.v(i) / 65536#)
        If i < 15 Then
            o.v(i + 1) = o.v(i + 1) + c - 1
        Else
            o.v(0) = o.v(0) + 38# * (c - 1)
        End If
        o.v(i) = o.v(i) - c * 65536#
    Next
End Sub

Private Sub sel25519(p As gf, q As gf, ByVal b As Long)
    Dim t As gf
    If b <> 0 Then
        t = p
        p = q
        q = t
    End If
End Sub

' a whole number in [0, 2^16): x - 65536 * floor(x / 65536), and C's (x >> 16) & 1 of a limb in [-2^16, 2^16)
Private Function mod65536(ByVal x As Double) As Double
    mod65536 = x - Int(x / 65536#) * 65536#
End Function

Private Sub pack25519(o() As Byte, ByVal oOfs As Long, n As gf)
    Dim i As Long, j As Long, b As Long, m As gf, t As gf, w As Long
    t = n
    car25519 t
    car25519 t
    car25519 t
    For j = 0 To 1
        m.v(0) = t.v(0) - 65517#                ' 0xffed
        For i = 1 To 14
            If m.v(i - 1) < 0 Then
                m.v(i) = t.v(i) - 65535# - 1
            Else
                m.v(i) = t.v(i) - 65535#
            End If
            m.v(i - 1) = mod65536(m.v(i - 1))
        Next
        If m.v(14) < 0 Then
            m.v(15) = t.v(15) - 32767# - 1
        Else
            m.v(15) = t.v(15) - 32767#
        End If
        If m.v(15) < 0 Then b = 1 Else b = 0
        m.v(14) = mod65536(m.v(14))
        If b = 0 Then t = m                     ' sel25519(t, m, 1 - b): t - p when t >= p
    Next
    For i = 0 To 15
        w = t.v(i)
        o(oOfs + 2 * i) = w And 255
        o(oOfs + 2 * i + 1) = w \ 256
    Next
End Sub

Private Function neq25519(a As gf, b As gf) As Boolean
    Dim c(0 To 31) As Byte, d(0 To 31) As Byte
    pack25519 c, 0, a
    pack25519 d, 0, b
    neq25519 = (crypto_verify_n(c, 0, d, 0, 32) <> 0)
End Function

Private Function par25519(a As gf) As Long
    Dim d(0 To 31) As Byte
    pack25519 d, 0, a
    par25519 = d(0) And 1
End Function

Private Function iszero25519(a As gf) As Boolean
    Dim d(0 To 31) As Byte, i As Long
    pack25519 d, 0, a
    For i = 0 To 31
        If d(i) <> 0 Then Exit Function
    Next
    iszero25519 = True
End Function

Private Sub unpack25519(o As gf, n() As Byte, ByVal nOfs As Long)
    Dim i As Long
    For i = 0 To 15
        o.v(i) = n(nOfs + 2 * i) + 256# * n(nOfs + 2 * i + 1)
    Next
    o.v(15) = CLng(o.v(15)) And &H7FFF&
End Sub

Private Sub fe_add(o As gf, a As gf, b As gf)       ' TweetNaCl's A
    Dim i As Long
    For i = 0 To 15
        o.v(i) = a.v(i) + b.v(i)
    Next
End Sub

Private Sub fe_sub(o As gf, a As gf, b As gf)       ' TweetNaCl's Z
    Dim i As Long
    For i = 0 To 15
        o.v(i) = a.v(i) - b.v(i)
    Next
End Sub

Private Sub fe_mul(o As gf, a As gf, b As gf)       ' TweetNaCl's M
    Dim i As Long, j As Long, ai As Double
    For i = 0 To 30
        mt(i) = 0
    Next
    For i = 0 To 15
        ai = a.v(i)
        For j = 0 To 15
            mt(i + j) = mt(i + j) + ai * b.v(j)
        Next
    Next
    For i = 0 To 14
        mt(i) = mt(i) + 38# * mt(i + 16)
    Next
#If SODIUM_LIMB_CHECK Then
    For i = 0 To 30
        limbSee mt(i)
    Next
    For i = 0 To 15
        limbSee a.v(i)
        limbSee b.v(i)
    Next
#End If
    For i = 0 To 15
        o.v(i) = mt(i)
    Next
    car25519 o
    car25519 o
End Sub

Private Sub fe_sq(o As gf, a As gf)                 ' TweetNaCl's S
    fe_mul o, a, a
End Sub

Private Sub inv25519(o As gf, inp As gf)
    Dim c As gf, a As Long
    c = inp
    For a = 253 To 0 Step -1
        fe_sq c, c
        If a <> 2 And a <> 4 Then fe_mul c, c, inp
    Next
    o = c
End Sub

Private Sub pow2523(o As gf, inp As gf)
    Dim c As gf, a As Long
    c = inp
    For a = 250 To 0 Step -1
        fe_sq c, c
        If a <> 1 Then fe_mul c, c, inp
    Next
    o = c
End Sub

Private Sub fe_small(o As gf, ByVal x As Long)      ' a small non-negative integer
    Dim i As Long
    For i = 0 To 15
        o.v(i) = 0
    Next
    o.v(0) = x And &HFFFF&
    o.v(1) = x \ 65536
End Sub

'--------------------------------------------------------------------------------------------------------
' the constants (computed, not copied: EDWARDS_D = -121665/121666, I = 2^((p-1)/4), the base point (x, 4/5) with x even)

Private Sub InitConstants()
    Dim t As gf, u As gf, e(0 To 31) As Byte, i As Long, k_ As Long
    Dim Ls As String
    For i = 0 To 7
        bitv(i) = 2 ^ i
    Next
    fe_small gf0, 0
    fe_small gf1, 1
    fe_small c121665, 121665
    ' EDWARDS_D
    fe_small t, 121666
    inv25519 u, t
    fe_small t, 121665
    fe_sub t, gf0, t
    fe_mul EDWARDS_D, t, u
    fe_add EDWARDS_D2, EDWARDS_D, EDWARDS_D
    car25519 EDWARDS_D2
    ' SQRTM1 = 2^((p - 1) / 4), the exponent 2^253 - 5 (bits 0, 1 and 3 .. 252 set)
    fe_small t, 2
    SQRTM1 = gf1
    For k_ = 252 To 0 Step -1
        fe_sq SQRTM1, SQRTM1
        If k_ <> 2 Then fe_mul SQRTM1, SQRTM1, t
    Next
    ' the base point: y = 4/5, x the even root
    fe_small t, 5
    inv25519 u, t
    fe_small t, 4
    fe_mul BASEY, t, u
    ready = True                                ' (unpack below needs the rest)
    pack25519 e, 0, BASEY
    Dim P As ge
    If Not ge_frombytes(P, e, 0, False) Then Err.Raise 5, "sodium_25519_c", "the base point does not decode"
    BASEX = P.X
    ' L = 2^252 + 27742317777372353535851937790883648493
    Ls = "edd3f55c1a631258d69cf7a2def9de1400000000000000000000000000000010"
    For i = 0 To 31
        Lbytes(i) = CByte("&H" & Mid$(Ls, 2 * i + 1, 2))
        Lc(i) = Lbytes(i)
    Next
End Sub

Private Sub EnsureReady()
    If Not ready Then InitConstants
End Sub

'--------------------------------------------------------------------------------------------------------
' X25519

' libsodium's x25519_ref10.c has_small_order: the 7 encodings of the curve's points of small order (the top bit
' ignored)
Private Function x25519_has_small_order(s() As Byte, ByVal sOfs As Long) As Boolean
    Static blacklist(0 To 6, 0 To 31) As Byte
    Static blReady As Boolean
    Dim rows() As String, i As Long, j As Long, c(0 To 6) As Long, k As Long
    If Not blReady Then
        rows = Split("0000000000000000000000000000000000000000000000000000000000000000 " & _
                     "0100000000000000000000000000000000000000000000000000000000000000 " & _
                     "e0eb7a7c3b41b8ae1656e3faf19fc46ada098deb9c32b1fd866205165f49b800 " & _
                     "5f9c95bca3508c24b1d0b1559c83ef5b04445cc4581c8e86d8224eddd09f1157 " & _
                     "ecffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f " & _
                     "edffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f " & _
                     "eeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f", " ")
        For i = 0 To 6
            For j = 0 To 31
                blacklist(i, j) = CByte("&H" & Mid$(rows(i), 2 * j + 1, 2))
            Next
        Next
        blReady = True
    End If
    For j = 0 To 30
        For i = 0 To 6
            c(i) = c(i) Or (s(sOfs + j) Xor blacklist(i, j))
        Next
    Next
    For i = 0 To 6
        c(i) = c(i) Or ((s(sOfs + 31) And &H7F) Xor blacklist(i, 31))
    Next
    For i = 0 To 6
        If c(i) = 0 Then k = 1
    Next
    x25519_has_small_order = (k = 1)
End Function

' crypto_scalarmult_curve25519(q, n, p): q = n * p (n clamped); -1 when p has small order or the result is 0
Public Function crypto_scalarmult_curve25519(q() As Byte, ByVal qOfs As Long, n() As Byte, ByVal nOfs As Long, p() As Byte, ByVal pOfs As Long) As Long
    Dim z(0 To 31) As Byte, i As Long, r As Long, x As gf
    Dim a As gf, b As gf, c As gf, d As gf, e As gf, f As gf
    EnsureReady
    If x25519_has_small_order(p, pOfs) Then
        crypto_scalarmult_curve25519 = -1
        Exit Function
    End If
    For i = 0 To 30
        z(i) = n(nOfs + i)
    Next
    z(31) = (n(nOfs + 31) And 127) Or 64
    z(0) = z(0) And 248
    unpack25519 x, p, pOfs
    b = x
    a = gf0
    c = gf0
    d = gf0
    a.v(0) = 1
    d.v(0) = 1
    For i = 254 To 0 Step -1
        r = (z(i \ 8) \ bitv(i And 7)) And 1
        sel25519 a, b, r
        sel25519 c, d, r
        fe_add e, a, c
        fe_sub a, a, c
        fe_add c, b, d
        fe_sub b, b, d
        fe_sq d, e
        fe_sq f, a
        fe_mul a, c, a
        fe_mul c, b, e
        fe_add e, a, c
        fe_sub a, a, c
        fe_sq b, a
        fe_sub c, d, f
        fe_mul a, c, c121665
        fe_add a, a, d
        fe_mul c, c, a
        fe_mul a, d, f
        fe_mul d, b, x
        fe_sq b, e
        sel25519 a, b, r
        sel25519 c, d, r
    Next
    inv25519 c, c
    fe_mul a, a, c
    pack25519 q, qOfs, a
    r = 0
    For i = 0 To 31
        r = r Or q(qOfs + i)
    Next
    If r = 0 Then crypto_scalarmult_curve25519 = -1 Else crypto_scalarmult_curve25519 = 0
End Function

'--------------------------------------------------------------------------------------------------------
' Ed25519's group (TweetNaCl's add, cswap, pack, scalarmult, scalarbase, unpackneg)

Private Sub ge_add(p As ge, q As ge)                ' p += q (q may be p)
    Dim a As gf, b As gf, c As gf, dd As gf, t As gf, e As gf, f As gf, g As gf, h As gf
    fe_sub a, p.Y, p.X
    fe_sub t, q.Y, q.X
    fe_mul a, a, t
    fe_add b, p.X, p.Y
    fe_add t, q.X, q.Y
    fe_mul b, b, t
    fe_mul c, p.T, q.T
    fe_mul c, c, EDWARDS_D2
    fe_mul dd, p.Z, q.Z
    fe_add dd, dd, dd
    fe_sub e, b, a
    fe_sub f, dd, c
    fe_add g, dd, c
    fe_add h, b, a
    fe_mul p.X, e, f
    fe_mul p.Y, h, g
    fe_mul p.Z, g, f
    fe_mul p.T, e, h
End Sub

Private Sub ge_cswap(p As ge, q As ge, ByVal b As Long)
    Dim t As ge
    If b <> 0 Then
        t = p
        p = q
        q = t
    End If
End Sub

Private Sub ge_pack(r() As Byte, ByVal rOfs As Long, p As ge)
    Dim tx As gf, ty As gf, zi As gf
    inv25519 zi, p.Z
    fe_mul tx, p.X, zi
    fe_mul ty, p.Y, zi
    pack25519 r, rOfs, ty
    r(rOfs + 31) = r(rOfs + 31) Xor (par25519(tx) * 128)
End Sub

Private Sub ge_identity(p As ge)
    p.X = gf0
    p.Y = gf1
    p.Z = gf1
    p.T = gf0
End Sub

' p = s * q, s 32 bytes little-endian (q is used up, as TweetNaCl's scalarmult does)
Private Sub ge_scalarmult(p As ge, q As ge, s() As Byte, ByVal sOfs As Long)
    Dim i As Long, b As Long
    ge_identity p
    For i = 255 To 0 Step -1
        b = (s(sOfs + i \ 8) \ bitv(i And 7)) And 1
        ge_cswap p, q, b
        ge_add q, p
        ge_add p, p
        ge_cswap p, q, b
    Next
End Sub

Private Sub ge_scalarbase(p As ge, s() As Byte, ByVal sOfs As Long)
    Dim q As ge
    q.X = BASEX
    q.Y = BASEY
    q.Z = gf1
    fe_mul q.T, BASEX, BASEY
    ge_scalarmult p, q, s, sOfs
End Sub

' ge25519_frombytes (negate False) and ge25519_frombytes_negate_vartime (negate True), as TweetNaCl's
' unpackneg: False when the encoding has no x (not a square)
Private Function ge_frombytes(r As ge, p() As Byte, ByVal pOfs As Long, ByVal negate As Boolean) As Boolean
    Dim t As gf, chk As gf, num As gf, den As gf, den2 As gf, den4 As gf, den6 As gf, s As Long
    r.Z = gf1
    unpack25519 r.Y, p, pOfs
    fe_sq num, r.Y
    fe_mul den, num, EDWARDS_D
    fe_sub num, num, r.Z
    fe_add den, r.Z, den

    fe_sq den2, den
    fe_sq den4, den2
    fe_mul den6, den4, den2
    fe_mul t, den6, num
    fe_mul t, t, den

    pow2523 t, t
    fe_mul t, t, num
    fe_mul t, t, den
    fe_mul t, t, den
    fe_mul r.X, t, den

    fe_sq chk, r.X
    fe_mul chk, chk, den
    If neq25519(chk, num) Then fe_mul r.X, r.X, SQRTM1

    fe_sq chk, r.X
    fe_mul chk, chk, den
    If neq25519(chk, num) Then Exit Function

    s = p(pOfs + 31) \ 128
    If negate Then
        If par25519(r.X) = s Then fe_sub r.X, gf0, r.X
    Else
        If par25519(r.X) <> s Then fe_sub r.X, gf0, r.X
    End If
    fe_mul r.T, r.X, r.Y
    ge_frombytes = True
End Function

' ge25519_has_small_order (a point): 8 * p is the identity (its X is 0)
Private Function ge_has_small_order(p As ge) As Boolean
    Dim q As ge
    q = p
    ge_add q, q
    ge_add q, q
    ge_add q, q
    ge_has_small_order = iszero25519(q.X)
End Function

' ge25519_is_on_main_subgroup: L * p is the identity
Private Function ge_is_on_main_subgroup(p As ge) As Boolean
    Dim q As ge, r As ge
    q = p
    ge_scalarmult r, q, Lbytes, 0
    ge_is_on_main_subgroup = iszero25519(r.X)
End Function

' ge25519_is_on_curve: (Y^2 - X^2) Z^2 = Z^4 + d X^2 Y^2
Private Function ge_is_on_curve(p As ge) As Boolean
    Dim x2 As gf, y2 As gf, z2 As gf, z4 As gf, t0 As gf, t1 As gf
    fe_sq x2, p.X
    fe_sq y2, p.Y
    fe_sq z2, p.Z
    fe_sub t0, y2, x2
    fe_mul t0, t0, z2
    fe_mul t1, x2, y2
    fe_mul t1, t1, EDWARDS_D
    fe_sq z4, z2
    fe_add t1, t1, z4
    fe_sub t0, t0, t1
    ge_is_on_curve = iszero25519(t0)
End Function

' ge25519_is_canonical(s): the encoded y is below p
Private Function ge25519_is_canonical(s() As Byte, ByVal sOfs As Long) As Boolean
    Dim c As Long, i As Long
    c = (s(sOfs + 31) And &H7F) Xor &H7F
    For i = 30 To 1 Step -1
        c = c Or (s(sOfs + i) Xor &HFF)
    Next
    ' c == 0 (y's top bytes all ones) and s[0] >= 0xed: not canonical
    ge25519_is_canonical = Not (c = 0 And s(sOfs) >= &HED)
End Function

' sc25519_is_canonical(s): s < L
Private Function sc25519_is_canonical(s() As Byte, ByVal sOfs As Long) As Boolean
    Dim i As Long
    EnsureReady
    For i = 31 To 0 Step -1
        If s(sOfs + i) < Lbytes(i) Then
            sc25519_is_canonical = True
            Exit Function
        ElseIf s(sOfs + i) > Lbytes(i) Then
            Exit Function
        End If
    Next
End Function                                ' s == L: not canonical

'--------------------------------------------------------------------------------------------------------
' scalars modulo L (TweetNaCl's modL and reduce, in Doubles)

Private Sub modL(r() As Byte, ByVal rOfs As Long, x() As Double)
    Dim carry As Double, i As Long, j As Long
    For i = 63 To 32 Step -1
        carry = 0
        For j = i - 32 To i - 13
            x(j) = x(j) + carry - 16# * x(i) * Lc(j - (i - 32))
#If SODIUM_LIMB_CHECK Then
            limbSee x(j)
            limbSee 16# * x(i) * Lc(j - (i - 32))
#End If
            carry = Int((x(j) + 128#) / 256#)
            x(j) = x(j) - carry * 256#
        Next
        x(i - 12) = x(i - 12) + carry
        x(i) = 0
    Next
    carry = 0
    For j = 0 To 31
        x(j) = x(j) + carry - Int(x(31) / 16#) * Lc(j)
        carry = Int(x(j) / 256#)
        x(j) = x(j) - carry * 256#
    Next
    For j = 0 To 31
        x(j) = x(j) - carry * Lc(j)
    Next
    For i = 0 To 31
        x(i + 1) = x(i + 1) + Int(x(i) / 256#)
        r(rOfs + i) = x(i) - Int(x(i) / 256#) * 256#
    Next
End Sub

' sc25519_reduce: the 64 bytes at r(rOfs) modulo L into r(rOfs .. rOfs + 31)
Private Sub sc_reduce(r() As Byte, ByVal rOfs As Long)
    Dim x(0 To 63) As Double, i As Long
    For i = 0 To 63
        x(i) = r(rOfs + i)
    Next
    For i = 0 To 63
        r(rOfs + i) = 0
    Next
    modL r, rOfs, x
End Sub

'--------------------------------------------------------------------------------------------------------
' Ed25519

' crypto_sign_ed25519_seed_keypair(pk, sk, seed): sk = seed || pk
Public Function crypto_sign_ed25519_seed_keypair(pk() As Byte, ByVal pkOfs As Long, sk() As Byte, ByVal skOfs As Long, seed() As Byte, ByVal seedOfs As Long) As Long
    Dim h(0 To 63) As Byte, A As ge, i As Long, s(0 To 31) As Byte
    EnsureReady
    For i = 0 To 31
        s(i) = seed(seedOfs + i)
    Next
    crypto_hash_sha512 h, 0, s, 0, 32
    h(0) = h(0) And 248
    h(31) = (h(31) And 127) Or 64
    ge_scalarbase A, h, 0
    ge_pack pk, pkOfs, A
    For i = 0 To 31
        sk(skOfs + i) = s(i)
        sk(skOfs + 32 + i) = pk(pkOfs + i)
    Next
End Function

' crypto_sign_ed25519_keypair(pk, sk): a random seed
Public Function crypto_sign_ed25519_keypair(pk() As Byte, ByVal pkOfs As Long, sk() As Byte, ByVal skOfs As Long) As Long
    Dim seed(0 To 31) As Byte, i As Long
    randombytes_buf seed, 0, 32
    crypto_sign_ed25519_keypair = crypto_sign_ed25519_seed_keypair(pk, pkOfs, sk, skOfs, seed, 0)
    For i = 0 To 31
        seed(i) = 0
    Next
End Function

' crypto_sign_ed25519_sk_to_pk(pk, sk): the public key stored in the secret key's second half
Public Function crypto_sign_ed25519_sk_to_pk(pk() As Byte, ByVal pkOfs As Long, sk() As Byte, ByVal skOfs As Long) As Long
    Dim i As Long
    For i = 0 To 31
        pk(pkOfs + i) = sk(skOfs + 32 + i)
    Next
End Function

' crypto_sign_ed25519_detached(sig, NULL, m, mlen, sk): the 64-byte signature R || S (deterministic)
Public Function crypto_sign_ed25519_detached(sig() As Byte, ByVal sigOfs As Long, m() As Byte, ByVal mOfs As Long, ByVal mlen As Long, sk() As Byte, ByVal skOfs As Long) As Long
    Dim az(0 To 63) As Byte, nonce(0 To 63) As Byte, hram(0 To 63) As Byte, st As crypto_hash_sha512_state
    Dim R As ge, x(0 To 63) As Double, i As Long, j As Long, rs(0 To 63) As Byte
    EnsureReady
    crypto_hash_sha512 az, 0, sk, skOfs, 32
    crypto_hash_sha512_init st
    crypto_hash_sha512_update st, az, 32, 32
    crypto_hash_sha512_update st, m, mOfs, mlen
    crypto_hash_sha512_final st, nonce, 0
    sc_reduce nonce, 0
    ge_scalarbase R, nonce, 0
    ge_pack rs, 0, R                            ' R || pk, hashed with the message
    For i = 0 To 31
        rs(32 + i) = sk(skOfs + 32 + i)
    Next
    crypto_hash_sha512_init st
    crypto_hash_sha512_update st, rs, 0, 64
    crypto_hash_sha512_update st, m, mOfs, mlen
    crypto_hash_sha512_final st, hram, 0
    sc_reduce hram, 0
    az(0) = az(0) And 248                       ' _crypto_sign_ed25519_clamp(az)
    az(31) = (az(31) And 127) Or 64
    For i = 0 To 31                             ' sc25519_muladd(sig + 32, hram, az, nonce)
        x(i) = nonce(i)
    Next
    For i = 0 To 31
        For j = 0 To 31
            x(i + j) = x(i + j) + CDbl(hram(i)) * az(j)
        Next
    Next
#If SODIUM_LIMB_CHECK Then
    For i = 0 To 63
        limbSee x(i)
    Next
#End If
    For i = 0 To 31
        sig(sigOfs + i) = rs(i)
    Next
    modL sig, sigOfs + 32, x
End Function

' crypto_sign_ed25519_verify_detached(sig, m, mlen, pk): 0 when the signature is good, -1 otherwise
Public Function crypto_sign_ed25519_verify_detached(sig() As Byte, ByVal sigOfs As Long, m() As Byte, ByVal mOfs As Long, ByVal mlen As Long, pk() As Byte, ByVal pkOfs As Long) As Long
    Dim A As ge, expectedR As ge, P As ge, Q As ge, h(0 To 63) As Byte, st As crypto_hash_sha512_state
    Dim pkb(0 To 31) As Byte, i As Long
    EnsureReady
    crypto_sign_ed25519_verify_detached = -1
    If (sig(sigOfs + 63) And 240) <> 0 Then
        If Not sc25519_is_canonical(sig, sigOfs + 32) Then Exit Function
    End If
    If Not ge25519_is_canonical(pk, pkOfs) Then Exit Function
    If Not ge_frombytes(A, pk, pkOfs, True) Then Exit Function
    If ge_has_small_order(A) Then Exit Function
    If Not ge_frombytes(expectedR, sig, sigOfs, False) Then Exit Function
    If ge_has_small_order(expectedR) Then Exit Function
    For i = 0 To 31
        pkb(i) = pk(pkOfs + i)
    Next
    crypto_hash_sha512_init st
    crypto_hash_sha512_update st, sig, sigOfs, 32
    crypto_hash_sha512_update st, pkb, 0, 32
    crypto_hash_sha512_update st, m, mOfs, mlen
    crypto_hash_sha512_final st, h, 0
    sc_reduce h, 0
    ge_scalarmult P, A, h, 0                    ' h * (-A)
    ge_scalarbase Q, sig, sigOfs + 32           ' s * B
    ge_add P, Q                                 ' sB - hA
    fe_sub P.X, gf0, P.X                        ' check = R - (sB - hA)
    fe_sub P.T, gf0, P.T
    ge_add expectedR, P
    If Not ge_has_small_order(expectedR) Then Exit Function
    crypto_sign_ed25519_verify_detached = 0
End Function

' crypto_core_ed25519_is_valid_point(p): 1 for the canonical encoding of a point of the main subgroup other than
' the identity, 0 otherwise
Public Function crypto_core_ed25519_is_valid_point(p() As Byte, ByVal pOfs As Long) As Long
    Dim P3 As ge
    EnsureReady
    If Not ge25519_is_canonical(p, pOfs) Then Exit Function
    If Not ge_frombytes(P3, p, pOfs, False) Then Exit Function
    If Not ge_is_on_curve(P3) Then Exit Function
    If Not ge_is_on_main_subgroup(P3) Then Exit Function
    If ge_has_small_order(P3) Then Exit Function
    crypto_core_ed25519_is_valid_point = 1
End Function

' crypto_sign_ed25519_pk_to_curve25519(curve25519_pk, ed25519_pk): u = (1 + y) / (1 - y); -1 for a point of small
' order, outside the main subgroup or not on the curve
Public Function crypto_sign_ed25519_pk_to_curve25519(cpk() As Byte, ByVal cOfs As Long, epk() As Byte, ByVal eOfs As Long) As Long
    Dim A As ge, x As gf, oneMinusY As gf
    EnsureReady
    crypto_sign_ed25519_pk_to_curve25519 = -1
    If Not ge_frombytes(A, epk, eOfs, True) Then Exit Function
    If ge_has_small_order(A) Then Exit Function
    If Not ge_is_on_main_subgroup(A) Then Exit Function
    fe_sub oneMinusY, gf1, A.Y
    fe_add x, gf1, A.Y
    inv25519 oneMinusY, oneMinusY
    fe_mul x, x, oneMinusY
    pack25519 cpk, cOfs, x
    crypto_sign_ed25519_pk_to_curve25519 = 0
End Function

' crypto_sign_ed25519_sk_to_curve25519(curve25519_sk, ed25519_sk): the clamped first half of SHA-512(seed)
Public Function crypto_sign_ed25519_sk_to_curve25519(csk() As Byte, ByVal cOfs As Long, esk() As Byte, ByVal eOfs As Long) As Long
    Dim h(0 To 63) As Byte, i As Long
    crypto_hash_sha512 h, 0, esk, eOfs, 32
    h(0) = h(0) And 248
    h(31) = (h(31) And 127) Or 64
    For i = 0 To 31
        csk(cOfs + i) = h(i)
    Next
End Function

'--------------------------------------------------------------------------------------------------------
' crypto_kx (kx/crypto_kx.c): q = the X25519 shared point; BLAKE2b-512(q || client_pk || server_pk) gives the
' client's rx (first half) and tx (second half), the server's tx and rx

Private Function kx_keys(keys() As Byte, sk() As Byte, ByVal skOfs As Long, otherPk() As Byte, ByVal otherOfs As Long, clientPk() As Byte, ByVal clientOfs As Long, serverPk() As Byte, ByVal serverOfs As Long) As Long
    Dim buf(0 To 95) As Byte, i As Long
    If crypto_scalarmult_curve25519(buf, 0, sk, skOfs, otherPk, otherOfs) <> 0 Then
        kx_keys = -1
        Exit Function
    End If
    For i = 0 To 31
        buf(32 + i) = clientPk(clientOfs + i)
        buf(64 + i) = serverPk(serverOfs + i)
    Next
    crypto_generichash_blake2b keys, 0, 64, buf, 0, 96
End Function

Public Function crypto_kx_client_session_keys(rx() As Byte, ByVal rxOfs As Long, tx() As Byte, ByVal txOfs As Long, client_pk() As Byte, ByVal cpOfs As Long, client_sk() As Byte, ByVal csOfs As Long, server_pk() As Byte, ByVal spOfs As Long) As Long
    Dim keys(0 To 63) As Byte, i As Long
    If kx_keys(keys, client_sk, csOfs, server_pk, spOfs, client_pk, cpOfs, server_pk, spOfs) <> 0 Then
        crypto_kx_client_session_keys = -1
        Exit Function
    End If
    For i = 0 To crypto_kx_SESSIONKEYBYTES - 1
        rx(rxOfs + i) = keys(i)
        tx(txOfs + i) = keys(i + crypto_kx_SESSIONKEYBYTES)
    Next
End Function

Public Function crypto_kx_server_session_keys(rx() As Byte, ByVal rxOfs As Long, tx() As Byte, ByVal txOfs As Long, server_pk() As Byte, ByVal spOfs As Long, server_sk() As Byte, ByVal ssOfs As Long, client_pk() As Byte, ByVal cpOfs As Long) As Long
    Dim keys(0 To 63) As Byte, i As Long
    If kx_keys(keys, server_sk, ssOfs, client_pk, cpOfs, client_pk, cpOfs, server_pk, spOfs) <> 0 Then
        crypto_kx_server_session_keys = -1
        Exit Function
    End If
    For i = 0 To crypto_kx_SESSIONKEYBYTES - 1
        tx(txOfs + i) = keys(i)
        rx(rxOfs + i) = keys(i + crypto_kx_SESSIONKEYBYTES)
    Next
End Function
