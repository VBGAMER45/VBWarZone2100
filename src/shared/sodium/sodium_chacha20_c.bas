Attribute VB_Name = "sodium_chacha20_c"
' (VB) libsodium's crypto_aead_xchacha20poly1305_ietf_encrypt and _decrypt
' (crypto_aead/xchacha20poly1305/aead_xchacha20poly1305.c) over what they call: crypto_core_hchacha20,
' crypto_stream_chacha20_ietf (RFC 8439's ChaCha20: a 32-bit block counter, a 96-bit nonce) and
' crypto_onetimeauth_poly1305 (TweetNaCl's byte-radix Poly1305: no product passes 2^31). crc.cpp's
' SessionKeys use them. Not 4.7.0's code: tools/crypto_test.py checks them against libsodium.
Option Explicit

Private Const crypto_aead_xchacha20poly1305_ietf_ABYTES As Long = 16

Private xs(0 To 15) As Long             ' (VB) the ChaCha state being permuted, and its input
Private xin(0 To 15) As Long

Private Sub qr(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long)
    xs(a) = xs(a) + xs(b): xs(d) = sodium_rotl32(xs(d) Xor xs(a), 16)
    xs(c) = xs(c) + xs(d): xs(b) = sodium_rotl32(xs(b) Xor xs(c), 12)
    xs(a) = xs(a) + xs(b): xs(d) = sodium_rotl32(xs(d) Xor xs(a), 8)
    xs(c) = xs(c) + xs(d): xs(b) = sodium_rotl32(xs(b) Xor xs(c), 7)
End Sub

' the 20 rounds over xs()
Private Sub chacha_rounds()
    Dim i As Long
    For i = 1 To 10
        qr 0, 4, 8, 12
        qr 1, 5, 9, 13
        qr 2, 6, 10, 14
        qr 3, 7, 11, 15
        qr 0, 5, 10, 15
        qr 1, 6, 11, 12
        qr 2, 7, 8, 13
        qr 3, 4, 9, 14
    Next
End Sub

' the constants "expand 32-byte k" and the key into xin(0 .. 11)
Private Sub chacha_keysetup(k() As Byte, ByVal kOfs As Long)
    Dim i As Long
    xin(0) = &H61707865
    xin(1) = &H3320646E
    xin(2) = &H79622D32
    xin(3) = &H6B206574
    For i = 0 To 7
        xin(4 + i) = sodium_load32_le(k, kOfs + 4 * i)
    Next
End Sub

' crypto_core_hchacha20(out, in, k, NULL): out(outOfs .. + 31) from the 16-byte input and the 32-byte key
Public Function crypto_core_hchacha20(out() As Byte, ByVal outOfs As Long, inBuf() As Byte, ByVal inOfs As Long, k() As Byte, ByVal kOfs As Long) As Long
    Dim i As Long
    chacha_keysetup k, kOfs
    For i = 0 To 3
        xin(12 + i) = sodium_load32_le(inBuf, inOfs + 4 * i)
    Next
    For i = 0 To 15
        xs(i) = xin(i)
    Next
    chacha_rounds
    For i = 0 To 3
        sodium_store32_le out, outOfs + 4 * i, xs(i)
        sodium_store32_le out, outOfs + 16 + 4 * i, xs(12 + i)
    Next
End Function

' crypto_stream_chacha20_ietf_xor_ic(c, m, mlen, n, ic, k): c = m xor the key stream from block ic on (m Nothing
' to xor: the key stream itself, crypto_stream_chacha20_ietf)
Public Function crypto_stream_chacha20_ietf_xor_ic(c() As Byte, ByVal cOfs As Long, m() As Byte, ByVal mOfs As Long, ByVal mlen As Long, ByVal useM As Boolean, n() As Byte, ByVal nOfs As Long, ByVal ic As Long, k() As Byte, ByVal kOfs As Long) As Long
    Dim i As Long, j As Long, blk(0 To 63) As Byte, take As Long
    chacha_keysetup k, kOfs
    xin(12) = ic
    xin(13) = sodium_load32_le(n, nOfs)
    xin(14) = sodium_load32_le(n, nOfs + 4)
    xin(15) = sodium_load32_le(n, nOfs + 8)
    j = 0
    Do While j < mlen
        For i = 0 To 15
            xs(i) = xin(i)
        Next
        chacha_rounds
        For i = 0 To 15
            sodium_store32_le blk, 4 * i, xs(i) + xin(i)
        Next
        take = mlen - j
        If take > 64 Then take = 64
        If useM Then
            For i = 0 To take - 1
                c(cOfs + j + i) = m(mOfs + j + i) Xor blk(i)
            Next
        Else
            For i = 0 To take - 1
                c(cOfs + j + i) = blk(i)
            Next
        End If
        j = j + take
        xin(12) = xin(12) + 1                   ' the 32-bit block counter
    Loop
End Function

' crypto_onetimeauth_poly1305(out, m, mlen, k) (TweetNaCl's crypto_onetimeauth)
Private Sub add1305(h() As Long, c() As Long)
    Dim j As Long, u As Long
    For j = 0 To 16
        u = u + h(j) + c(j)
        h(j) = u And 255
        u = u \ 256
    Next
End Sub

Public Function crypto_onetimeauth_poly1305(out() As Byte, ByVal outOfs As Long, m() As Byte, ByVal mOfs As Long, ByVal n As Long, k() As Byte, ByVal kOfs As Long) As Long
    Dim s As Long, i As Long, j As Long, u As Long
    Dim x(0 To 16) As Long, r(0 To 16) As Long, h(0 To 16) As Long, c(0 To 16) As Long, g(0 To 16) As Long
    Dim minusp(0 To 16) As Long
    minusp(0) = 5
    minusp(16) = 252
    For j = 0 To 15
        r(j) = k(kOfs + j)
    Next
    r(3) = r(3) And 15
    r(4) = r(4) And 252
    r(7) = r(7) And 15
    r(8) = r(8) And 252
    r(11) = r(11) And 15
    r(12) = r(12) And 252
    r(15) = r(15) And 15

    Do While n > 0
        For j = 0 To 16
            c(j) = 0
        Next
        j = 0
        Do While j < 16 And j < n
            c(j) = m(mOfs + j)
            j = j + 1
        Loop
        c(j) = 1
        mOfs = mOfs + j
        n = n - j
        add1305 h, c
        For i = 0 To 16
            x(i) = 0
            For j = 0 To 16
                If j <= i Then
                    x(i) = x(i) + h(j) * r(i - j)
                Else
                    x(i) = x(i) + h(j) * (320 * r(i + 17 - j))
                End If
            Next
        Next
        For i = 0 To 16
            h(i) = x(i)
        Next
        u = 0
        For j = 0 To 15
            u = u + h(j)
            h(j) = u And 255
            u = u \ 256
        Next
        u = u + h(16)
        h(16) = u And 3
        u = 5 * (u \ 4)
        For j = 0 To 15
            u = u + h(j)
            h(j) = u And 255
            u = u \ 256
        Next
        u = u + h(16)
        h(16) = u
    Loop

    For j = 0 To 16
        g(j) = h(j)
    Next
    add1305 h, minusp
    s = -(h(16) \ 128)                          ' s = -(h[16] >> 7): all ones when h - p went negative
    For j = 0 To 16
        h(j) = h(j) Xor (s And (g(j) Xor h(j)))
    Next

    For j = 0 To 15
        c(j) = k(kOfs + j + 16)
    Next
    c(16) = 0
    add1305 h, c
    For j = 0 To 15
        out(outOfs + j) = h(j) And 255
    Next
End Function

' the Poly1305 input of the IETF construction: ad, its padding to 16, the ciphertext, its padding, the two
' lengths as 64-bit little-endian; the tag into tag(tagOfs .. + 15)
Private Sub aead_mac(tag() As Byte, ByVal tagOfs As Long, c() As Byte, ByVal cOfs As Long, ByVal clen As Long, ad() As Byte, ByVal adOfs As Long, ByVal adlen As Long, block0() As Byte)
    Dim buf() As Byte, n As Long, p As Long, i As Long, hi As Long, lo As Long
    n = adlen + ((16 - (adlen And 15)) And 15) + clen + ((16 - (clen And 15)) And 15) + 16
    ReDim buf(0 To n - 1)
    For i = 0 To adlen - 1
        buf(i) = ad(adOfs + i)
    Next
    p = adlen + ((16 - (adlen And 15)) And 15)
    For i = 0 To clen - 1
        buf(p + i) = c(cOfs + i)
    Next
    p = p + clen + ((16 - (clen And 15)) And 15)
    sodium_split64 adlen, hi, lo
    sodium_store32_le buf, p, lo
    sodium_store32_le buf, p + 4, hi
    sodium_split64 clen, hi, lo
    sodium_store32_le buf, p + 8, lo
    sodium_store32_le buf, p + 12, hi
    crypto_onetimeauth_poly1305 tag, tagOfs, buf, 0, n, block0, 0
End Sub

' the XChaCha20 subkey and IETF nonce: k2 = HChaCha20(k, npub[0..16]), npub2 = 4 zero bytes, npub[16..24]
Private Sub xchacha_subkey(k2() As Byte, npub2() As Byte, npub() As Byte, ByVal npubOfs As Long, k() As Byte, ByVal kOfs As Long)
    Dim i As Long
    crypto_core_hchacha20 k2, 0, npub, npubOfs, k, kOfs
    For i = 0 To 3
        npub2(i) = 0
    Next
    For i = 0 To 7
        npub2(4 + i) = npub(npubOfs + 16 + i)
    Next
End Sub

' crypto_aead_xchacha20poly1305_ietf_encrypt(c, &clen, m, mlen, ad, adlen, NULL, npub, k): the ciphertext and
' its 16-byte tag into c(cOfs .. cOfs + mlen + 15); returns 0 (clen = mlen + 16)
Public Function crypto_aead_xchacha20poly1305_ietf_encrypt(c() As Byte, ByVal cOfs As Long, clen As Long, m() As Byte, ByVal mOfs As Long, ByVal mlen As Long, ad() As Byte, ByVal adOfs As Long, ByVal adlen As Long, npub() As Byte, ByVal npubOfs As Long, k() As Byte, ByVal kOfs As Long) As Long
    Dim k2(0 To 31) As Byte, npub2(0 To 11) As Byte, block0(0 To 63) As Byte, dummy() As Byte
    xchacha_subkey k2, npub2, npub, npubOfs, k, kOfs
    crypto_stream_chacha20_ietf_xor_ic block0, 0, dummy, 0, 64, False, npub2, 0, 0, k2, 0
    crypto_stream_chacha20_ietf_xor_ic c, cOfs, m, mOfs, mlen, True, npub2, 0, 1, k2, 0
    aead_mac c, cOfs + mlen, c, cOfs, mlen, ad, adOfs, adlen, block0
    clen = mlen + crypto_aead_xchacha20poly1305_ietf_ABYTES
    crypto_aead_xchacha20poly1305_ietf_encrypt = 0
End Function

' crypto_aead_xchacha20poly1305_ietf_decrypt(m, &mlen, NULL, c, clen, ad, adlen, npub, k): 0 and the message
' in m(mOfs .. mOfs + mlen - 1), or -1 (too short, or forged: m zeroed, mlen 0)
Public Function crypto_aead_xchacha20poly1305_ietf_decrypt(m() As Byte, ByVal mOfs As Long, mlen As Long, c() As Byte, ByVal cOfs As Long, ByVal clen As Long, ad() As Byte, ByVal adOfs As Long, ByVal adlen As Long, npub() As Byte, ByVal npubOfs As Long, k() As Byte, ByVal kOfs As Long) As Long
    Dim k2(0 To 31) As Byte, npub2(0 To 11) As Byte, block0(0 To 63) As Byte, dummy() As Byte, tag(0 To 15) As Byte
    Dim n As Long, i As Long
    mlen = 0
    crypto_aead_xchacha20poly1305_ietf_decrypt = -1
    If clen < crypto_aead_xchacha20poly1305_ietf_ABYTES Then Exit Function
    n = clen - crypto_aead_xchacha20poly1305_ietf_ABYTES
    xchacha_subkey k2, npub2, npub, npubOfs, k, kOfs
    crypto_stream_chacha20_ietf_xor_ic block0, 0, dummy, 0, 64, False, npub2, 0, 0, k2, 0
    aead_mac tag, 0, c, cOfs, n, ad, adOfs, adlen, block0
    If crypto_verify_n(tag, 0, c, cOfs + n, 16) <> 0 Then
        For i = 0 To n - 1
            m(mOfs + i) = 0
        Next
        Exit Function
    End If
    crypto_stream_chacha20_ietf_xor_ic m, mOfs, c, cOfs, n, True, npub2, 0, 1, k2, 0
    mlen = n
    crypto_aead_xchacha20poly1305_ietf_decrypt = 0
End Function
