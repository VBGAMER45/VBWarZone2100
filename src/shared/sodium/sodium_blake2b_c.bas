Attribute VB_Name = "sodium_blake2b_c"
' (VB) libsodium's crypto_generichash, BLAKE2b (RFC 7693; crypto_generichash/blake2b/ref/blake2b-ref.c),
' unkeyed, in one call: crypto_kx's session keys (sodium_25519_c). A uint64_t word is a pair of Longs
' (sodium_utils_c). Not 4.7.0's code: tools/crypto_test.py checks it against libsodium.
Option Explicit

Private Const BLAKE2B_BLOCKBYTES As Long = 128

Private sigma(0 To 11, 0 To 15) As Long
Private vh(0 To 15) As Long
Private vl(0 To 15) As Long
Private mh(0 To 15) As Long
Private ml(0 To 15) As Long
Private sigmaReady As Boolean

Private Sub InitSigma()
    Dim s() As String, r As Long, i As Long
    s = Split("0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 " & _
              "14 10 4 8 9 15 13 6 1 12 0 2 11 7 5 3 " & _
              "11 8 12 0 5 2 15 13 10 14 3 6 7 1 9 4 " & _
              "7 9 3 1 13 12 11 14 2 6 5 10 4 0 15 8 " & _
              "9 0 5 7 2 4 10 15 14 1 11 12 6 8 3 13 " & _
              "2 12 6 10 0 11 8 3 4 13 7 5 15 14 1 9 " & _
              "12 5 1 15 14 13 4 10 0 7 6 3 9 2 8 11 " & _
              "13 11 7 14 12 1 3 9 5 0 15 4 8 6 2 10 " & _
              "6 15 14 9 11 3 0 8 12 2 13 7 1 4 10 5 " & _
              "10 2 8 4 7 6 1 5 15 11 9 14 3 12 13 0", " ")
    For r = 0 To 9
        For i = 0 To 15
            sigma(r, i) = CLng(s(r * 16 + i))
        Next
    Next
    For i = 0 To 15                         ' rounds 10 and 11 are rounds 0 and 1 again
        sigma(10, i) = sigma(0, i)
        sigma(11, i) = sigma(1, i)
    Next
    sigmaReady = True
End Sub

' G(r, i, a, b, c, d)
Private Sub blake2b_G(ByVal r As Long, ByVal i As Long, ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long)
    Dim x As Long
    x = sigma(r, 2 * i)
    sodium_add64 vh(a), vl(a), vh(a), vl(a), vh(b), vl(b)
    sodium_add64 vh(a), vl(a), vh(a), vl(a), mh(x), ml(x)
    sodium_rotr64 vh(d), vl(d), vh(d) Xor vh(a), vl(d) Xor vl(a), 32
    sodium_add64 vh(c), vl(c), vh(c), vl(c), vh(d), vl(d)
    sodium_rotr64 vh(b), vl(b), vh(b) Xor vh(c), vl(b) Xor vl(c), 24
    x = sigma(r, 2 * i + 1)
    sodium_add64 vh(a), vl(a), vh(a), vl(a), vh(b), vl(b)
    sodium_add64 vh(a), vl(a), vh(a), vl(a), mh(x), ml(x)
    sodium_rotr64 vh(d), vl(d), vh(d) Xor vh(a), vl(d) Xor vl(a), 16
    sodium_add64 vh(c), vl(c), vh(c), vl(c), vh(d), vl(d)
    sodium_rotr64 vh(b), vl(b), vh(b) Xor vh(c), vl(b) Xor vl(c), 63
End Sub

' blake2b_compress: h = the chain value (hh, hl), the block at blk(ofs), t the bytes so far, last the final block
Private Sub blake2b_compress(hh() As Long, hl() As Long, blk() As Byte, ByVal ofs As Long, ByVal t As Double, ByVal last As Boolean)
    Dim i As Long, r As Long, th As Long, tl As Long
    For i = 0 To 15
        mh(i) = sodium_load32_le(blk, ofs + 8 * i + 4)
        ml(i) = sodium_load32_le(blk, ofs + 8 * i)
    Next
    For i = 0 To 7
        vh(i) = hh(i)
        vl(i) = hl(i)
        sodium_sha512_iv i, vh(i + 8), vl(i + 8)
    Next
    sodium_split64 t, th, tl                 ' t[0]; t[1] (the counter's high 64 bits) is 0
    vh(12) = vh(12) Xor th
    vl(12) = vl(12) Xor tl
    If last Then
        vh(14) = Not vh(14)
        vl(14) = Not vl(14)
    End If
    For r = 0 To 11
        blake2b_G r, 0, 0, 4, 8, 12
        blake2b_G r, 1, 1, 5, 9, 13
        blake2b_G r, 2, 2, 6, 10, 14
        blake2b_G r, 3, 3, 7, 11, 15
        blake2b_G r, 4, 0, 5, 10, 15
        blake2b_G r, 5, 1, 6, 11, 12
        blake2b_G r, 6, 2, 7, 8, 13
        blake2b_G r, 7, 3, 4, 9, 14
    Next
    For i = 0 To 7
        hh(i) = hh(i) Xor vh(i) Xor vh(i + 8)
        hl(i) = hl(i) Xor vl(i) Xor vl(i + 8)
    Next
End Sub

' crypto_generichash(out, outlen, in, inlen, NULL, 0): BLAKE2b with an outlen-byte digest (1 .. 64), no key
Public Function crypto_generichash_blake2b(out() As Byte, ByVal outOfs As Long, ByVal outlen As Long, inBuf() As Byte, ByVal ofs As Long, ByVal inlen As Long) As Long
    Dim hh(0 To 7) As Long, hl(0 To 7) As Long, i As Long, t As Double, blk(0 To 127) As Byte, n As Long
    Dim b(0 To 63) As Byte
    If outlen < 1 Or outlen > 64 Then
        crypto_generichash_blake2b = -1
        Exit Function
    End If
    If Not sigmaReady Then InitSigma
    For i = 0 To 7
        sodium_sha512_iv i, hh(i), hl(i)
    Next
    hl(0) = hl(0) Xor (&H1010000 Or outlen)  ' the parameter block: digest length, key length 0, fanout 1, depth 1
    Do While inlen > BLAKE2B_BLOCKBYTES     ' every block but the last
        t = t + BLAKE2B_BLOCKBYTES
        blake2b_compress hh, hl, inBuf, ofs, t, False
        ofs = ofs + BLAKE2B_BLOCKBYTES
        inlen = inlen - BLAKE2B_BLOCKBYTES
    Loop
    For i = 0 To inlen - 1                  ' the last block (empty when the input is), padded with zeros
        blk(i) = inBuf(ofs + i)
    Next
    t = t + inlen
    blake2b_compress hh, hl, blk, 0, t, True
    For i = 0 To 7
        sodium_store32_le b, 8 * i, hl(i)
        sodium_store32_le b, 8 * i + 4, hh(i)
    Next
    For i = 0 To outlen - 1
        out(outOfs + i) = b(i)
    Next
End Function
