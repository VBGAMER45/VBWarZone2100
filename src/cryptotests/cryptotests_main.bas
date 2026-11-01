Attribute VB_Name = "cryptotests_main"
' (VB) CryptoTests.exe: tools/crypto_test.py's driver of src/shared/sodium (compiled in without WzFramework).
' CryptoTests.exe <commands file> <results file>: each line of the commands file is a command and its
' tab-separated arguments (bytes in hex, "-" for none); each writes one result line (tab-separated, hex),
' after any LOG / ASSERT lines the command made:
'   sha256 <m>   sha512 <m>   sha256chunks <m> <chunk>   sha512chunks <m> <chunk>   blake2b <outlen> <m>
'   hchacha20 <k> <in16>   chacha20 <k> <n12> <ic> <m>   poly1305 <k> <m>
'   aead_enc <k> <npub> <ad> <m>   aead_dec <k> <npub> <ad> <c>   x25519 <n> <p>
'   ed_seedkp <seed>   ed_sign <sk> <m>   ed_verify <sig> <m> <pk>   ed_valid <p>   ed_pk2c <pk>   ed_sk2c <sk>
'   kx_client <cpk> <csk> <spk>   kx_server <spk> <ssk> <cpk>   pad <buf> <len> <blocksize> <max>
'   unpad <buf> <blocksize>   sha256sum <m>   eckey_gen   eckey_pub <pk>   eckey_priv <sk>   eckey_empty
'   eckey_sign <sk> <m>   eckey_verify <pk> <sig> <m>   sk_init <mesk> <meidx> <otherpk> <otheridx>
'   sk_enc <mesk> <meidx> <otherpk> <otheridx> <m>   sk_dec <mesk> <meidx> <otherpk> <otheridx> <data>
'   b64enc <m>   b64dec <text>   b64url <text>   b64std <text>   random <n>   speed <rounds>
'   limbmax (sodium_25519_c's largest magnitude so far: the project sets SODIUM_LIMB_CHECK)
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Private outLines() As String
Private outCount As Long
Private freq As Currency

Private Sub emit(ByVal s As String)
    If outCount > UBound(outLines) Then ReDim Preserve outLines(0 To 2 * outCount)
    outLines(outCount) = s
    outCount = outCount + 1
End Sub

Public Sub testLog(ByVal s As String)
    emit s
End Sub

Private Function nowMs() As Double
    Dim c As Currency
    QueryPerformanceCounter c
    nowMs = CDbl(c) * 1000# / CDbl(freq)
End Function

' hex -> bytes (returns the count; "-" or "" is none)
Private Function unhex(ByVal s As String, b() As Byte) As Long
    Dim n As Long, i As Long
    If s = "-" Or s = "" Then
        Erase b
        unhex = 0
        Exit Function
    End If
    n = Len(s) \ 2
    ReDim b(0 To n - 1)
    For i = 0 To n - 1
        b(i) = CByte("&H" & Mid$(s, 2 * i + 1, 2))
    Next
    unhex = n
End Function

Private Function tohex(b() As Byte, ByVal ofs As Long, ByVal n As Long) As String
    Dim s As String, i As Long
    If n <= 0 Then
        tohex = "-"
        Exit Function
    End If
    s = Space$(2 * n)
    For i = 0 To n - 1
        Mid$(s, 2 * i + 1, 2) = Right$("0" & LCase$(Hex$(b(ofs + i))), 2)
    Next
    tohex = s
End Function

Private Function ecKeyFromSk(k As EcKey, ByVal hexSk As String) As Boolean
    Dim b() As Byte, n As Long
    n = unhex(hexSk, b)
    ecKeyFromSk = EcKey_fromBytes(k, b, n, EcKey_Private)
End Function

Private Function ecKeyFromPk(k As EcKey, ByVal hexPk As String) As Boolean
    Dim b() As Byte, n As Long
    n = unhex(hexPk, b)
    ecKeyFromPk = EcKey_fromBytes(k, b, n, EcKey_Public)
End Function

Private Sub runCommand(ByVal cmdLine As String)
    Dim f() As String, a() As Byte, b() As Byte, c() As Byte, d() As Byte, o() As Byte, o2() As Byte
    Dim na As Long, nb As Long, nc As Long, nd As Long, n As Long, r As Long, i As Long, chunk As Long
    Dim st256 As crypto_hash_sha256_state, st512 As crypto_hash_sha512_state, h As Sha256
    Dim k As EcKey, k2 As EcKey, sk As SessionKeys, ok As Boolean, s As String
    f = Split(cmdLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab, vbTab)
    Select Case f(0)
    Case "sha256"
        na = unhex(f(1), a)
        ReDim o(0 To 31)
        crypto_hash_sha256 o, 0, a, 0, na
        emit tohex(o, 0, 32)
    Case "sha512"
        na = unhex(f(1), a)
        ReDim o(0 To 63)
        crypto_hash_sha512 o, 0, a, 0, na
        emit tohex(o, 0, 64)
    Case "sha256chunks", "sha512chunks"
        na = unhex(f(1), a)
        chunk = CLng(f(2))
        ReDim o(0 To 63)
        If f(0) = "sha256chunks" Then crypto_hash_sha256_init st256 Else crypto_hash_sha512_init st512
        i = 0
        Do While i < na
            n = chunk
            If i + n > na Then n = na - i
            If f(0) = "sha256chunks" Then crypto_hash_sha256_update st256, a, i, n Else crypto_hash_sha512_update st512, a, i, n
            i = i + n
        Loop
        If f(0) = "sha256chunks" Then
            crypto_hash_sha256_final st256, o, 0
            emit tohex(o, 0, 32)
        Else
            crypto_hash_sha512_final st512, o, 0
            emit tohex(o, 0, 64)
        End If
    Case "blake2b"
        n = CLng(f(1))
        na = unhex(f(2), a)
        ReDim o(0 To 63)
        r = crypto_generichash_blake2b(o, 0, n, a, 0, na)
        If r <> 0 Then emit "FAIL" Else emit tohex(o, 0, n)
    Case "hchacha20"
        na = unhex(f(1), a)
        nb = unhex(f(2), b)
        ReDim o(0 To 31)
        crypto_core_hchacha20 o, 0, b, 0, a, 0
        emit tohex(o, 0, 32)
    Case "chacha20"
        na = unhex(f(1), a)
        nb = unhex(f(2), b)
        nd = unhex(f(4), d)
        If nd > 0 Then ReDim o(0 To nd - 1)
        crypto_stream_chacha20_ietf_xor_ic o, 0, d, 0, nd, True, b, 0, CLng(f(3)), a, 0
        emit tohex(o, 0, nd)
    Case "poly1305"
        na = unhex(f(1), a)
        nb = unhex(f(2), b)
        ReDim o(0 To 15)
        crypto_onetimeauth_poly1305 o, 0, b, 0, nb, a, 0
        emit tohex(o, 0, 16)
    Case "aead_enc"
        na = unhex(f(1), a)
        nb = unhex(f(2), b)
        nc = unhex(f(3), c)
        nd = unhex(f(4), d)
        ReDim o(0 To nd + 15)
        crypto_aead_xchacha20poly1305_ietf_encrypt o, 0, n, d, 0, nd, c, 0, nc, b, 0, a, 0
        emit tohex(o, 0, n)
    Case "aead_dec"
        na = unhex(f(1), a)
        nb = unhex(f(2), b)
        nc = unhex(f(3), c)
        nd = unhex(f(4), d)
        ReDim o(0 To nd + 1)
        r = crypto_aead_xchacha20poly1305_ietf_decrypt(o, 0, n, d, 0, nd, c, 0, nc, b, 0, a, 0)
        If r <> 0 Then emit "FAIL" Else emit tohex(o, 0, n)
    Case "x25519"
        na = unhex(f(1), a)
        nb = unhex(f(2), b)
        ReDim o(0 To 31)
        r = crypto_scalarmult_curve25519(o, 0, a, 0, b, 0)
        If r <> 0 Then emit "FAIL" Else emit tohex(o, 0, 32)
    Case "ed_seedkp"
        na = unhex(f(1), a)
        ReDim o(0 To 31)
        ReDim o2(0 To 63)
        crypto_sign_ed25519_seed_keypair o, 0, o2, 0, a, 0
        emit tohex(o, 0, 32) & vbTab & tohex(o2, 0, 64)
    Case "ed_sign"
        na = unhex(f(1), a)
        nb = unhex(f(2), b)
        ReDim o(0 To 63)
        crypto_sign_ed25519_detached o, 0, b, 0, nb, a, 0
        emit tohex(o, 0, 64)
    Case "ed_verify"
        na = unhex(f(1), a)
        nb = unhex(f(2), b)
        nc = unhex(f(3), c)
        emit CStr(crypto_sign_ed25519_verify_detached(a, 0, b, 0, nb, c, 0))
    Case "ed_valid"
        na = unhex(f(1), a)
        emit CStr(crypto_core_ed25519_is_valid_point(a, 0))
    Case "ed_pk2c"
        na = unhex(f(1), a)
        ReDim o(0 To 31)
        r = crypto_sign_ed25519_pk_to_curve25519(o, 0, a, 0)
        If r <> 0 Then emit "FAIL" Else emit tohex(o, 0, 32)
    Case "ed_sk2c"
        na = unhex(f(1), a)
        ReDim o(0 To 31)
        crypto_sign_ed25519_sk_to_curve25519 o, 0, a, 0
        emit tohex(o, 0, 32)
    Case "kx_client", "kx_server"
        na = unhex(f(1), a)
        nb = unhex(f(2), b)
        nc = unhex(f(3), c)
        ReDim o(0 To 31)
        ReDim o2(0 To 31)
        If f(0) = "kx_client" Then
            r = crypto_kx_client_session_keys(o, 0, o2, 0, a, 0, b, 0, c, 0)
        Else
            r = crypto_kx_server_session_keys(o, 0, o2, 0, a, 0, b, 0, c, 0)
        End If
        If r <> 0 Then emit "FAIL" Else emit tohex(o, 0, 32) & vbTab & tohex(o2, 0, 32)
    Case "pad"
        na = unhex(f(1), a)
        r = sodium_pad(n, a, CLng(f(2)), CLng(f(3)), CLng(f(4)))
        emit r & vbTab & n & vbTab & tohex(a, 0, na)
    Case "unpad"
        na = unhex(f(1), a)
        r = sodium_unpad(n, a, na, CLng(f(2)))
        emit r & vbTab & n
    Case "sha256sum"
        na = unhex(f(1), a)
        h = sha256Sum(a, 0, na)
        emit Sha256_toString(h)
    Case "eckey_gen"
        EcKey_generate k
        n = EcKey_toBytes(k, EcKey_Private, o)
        nb = EcKey_toBytes(k, EcKey_Public, o2)
        emit tohex(o, 0, n) & vbTab & tohex(o2, 0, nb) & vbTab & EcKey_hasPrivate(k)
    Case "eckey_pub"
        ok = ecKeyFromPk(k, f(1))
        n = EcKey_toBytes(k, EcKey_Public, o)
        nb = EcKey_toBytes(k, EcKey_Private, o2)
        emit ok & vbTab & tohex(o, 0, n) & vbTab & nb & vbTab & EcKey_publicHashString(k) & vbTab & _
             EcKey_publicHashString(k, 10) & vbTab & EcKey_publicKeyHexString(k, 8) & vbTab & EcKey_hasPrivate(k)
    Case "eckey_priv"
        ok = ecKeyFromSk(k, f(1))
        n = EcKey_toBytes(k, EcKey_Public, o)
        nb = EcKey_toBytes(k, EcKey_Private, o2)
        emit ok & vbTab & tohex(o, 0, n) & vbTab & tohex(o2, 0, nb) & vbTab & EcKey_hasPrivate(k)
    Case "eckey_empty"
        emit EcKey_empty(k) & vbTab & EcKey_hasPrivate(k) & vbTab & EcKey_publicHashString(k) & vbTab & _
             EcKey_publicKeyHexString(k) & vbTab & EcKey_sign(k, a, 0, 0, o) & vbTab & EcKey_verify(k, a, 0, b, 0, 0)
    Case "eckey_sign"
        ok = ecKeyFromSk(k, f(1))
        nb = unhex(f(2), b)
        n = EcKey_sign(k, b, 0, nb, o)
        emit tohex(o, 0, n)
    Case "eckey_verify"
        ok = ecKeyFromPk(k, f(1))
        na = unhex(f(2), a)
        nb = unhex(f(3), b)
        emit EcKey_verify(k, a, na, b, 0, nb)
    Case "sk_init", "sk_enc", "sk_dec"
        ok = ecKeyFromSk(k, f(1))
        ok = ecKeyFromPk(k2, f(3))
        ok = SessionKeys_init(sk, k, CLng(f(2)), k2, CLng(f(4)))
        If f(0) = "sk_init" Then
            If ok Then
                emit "1" & vbTab & tohex(sk.receiveKey, 0, sk.receiveKeySize) & vbTab & tohex(sk.sendKey, 0, sk.sendKeySize)
            Else
                emit "0" & vbTab & SessionKeys_lastError()
            End If
        ElseIf f(0) = "sk_enc" Then
            nd = unhex(f(5), d)
            n = SessionKeys_encryptMessageForOther(sk, d, 0, nd, o)
            emit tohex(o, 0, n)
        Else
            nd = unhex(f(5), d)
            ok = SessionKeys_decryptMessageFromOther(sk, d, 0, nd, o, n)
            If ok Then emit "1" & vbTab & tohex(o, 0, n) Else emit "0"
        End If
    Case "b64enc"
        na = unhex(f(1), a)
        emit base64Encode(a, na)
    Case "b64dec"
        n = base64Decode(f(1), o)
        emit tohex(o, 0, n)
    Case "b64url"
        emit b64Tob64UrlSafe(f(1))
    Case "b64std"
        emit b64UrlSafeTob64(f(1))
    Case "random"
        n = CLng(f(1))
        ReDim o(0 To n - 1)
        genSecRandomBytes o, 0, n
        emit tohex(o, 0, n)
    Case "speed"
        runSpeed CLng(f(1))
    Case "limbmax"
        emit CStr(sodium_25519_limbMax())
    Case Else
        emit "UNKNOWN " & f(0)
    End Select
End Sub

' times per operation, in milliseconds
Private Sub runSpeed(ByVal rounds As Long)
    Dim t As Double, i As Long, big() As Byte, o(0 To 63) As Byte, sk(0 To 63) As Byte, pk(0 To 31) As Byte
    Dim sig(0 To 63) As Byte, msg(0 To 99) As Byte, q(0 To 31) As Byte, seed(0 To 31) As Byte, n As Long
    Dim c(0 To 1039) As Byte, kk(0 To 31) As Byte, npub(0 To 23) As Byte, noAd() As Byte, s As String
    ReDim big(0 To 1048575)
    t = nowMs()
    crypto_hash_sha256 o, 0, big, 0, 1048576
    s = "sha256 1MiB " & Format$(nowMs() - t, "0.0")
    t = nowMs()
    crypto_hash_sha512 o, 0, big, 0, 1048576
    s = s & vbTab & "sha512 1MiB " & Format$(nowMs() - t, "0.0")
    t = nowMs()
    For i = 1 To rounds
        crypto_sign_ed25519_seed_keypair pk, 0, sk, 0, seed, 0
    Next
    s = s & vbTab & "keypair " & Format$((nowMs() - t) / rounds, "0.0")
    t = nowMs()
    For i = 1 To rounds
        crypto_sign_ed25519_detached sig, 0, msg, 0, 100, sk, 0
    Next
    s = s & vbTab & "sign " & Format$((nowMs() - t) / rounds, "0.0")
    t = nowMs()
    For i = 1 To rounds
        n = crypto_sign_ed25519_verify_detached(sig, 0, msg, 0, 100, pk, 0)
    Next
    s = s & vbTab & "verify " & Format$((nowMs() - t) / rounds, "0.0") & " (" & n & ")"
    t = nowMs()
    For i = 1 To rounds
        n = crypto_core_ed25519_is_valid_point(pk, 0)
    Next
    s = s & vbTab & "is_valid_point " & Format$((nowMs() - t) / rounds, "0.0")
    t = nowMs()
    For i = 1 To rounds
        crypto_scalarmult_curve25519 q, 0, sk, 0, pk, 0
    Next
    s = s & vbTab & "x25519 " & Format$((nowMs() - t) / rounds, "0.0")
    t = nowMs()
    For i = 1 To rounds
        crypto_aead_xchacha20poly1305_ietf_encrypt c, 0, n, big, 0, 1024, noAd, 0, 0, npub, 0, kk, 0
    Next
    s = s & vbTab & "aead 1KiB " & Format$((nowMs() - t) / rounds, "0.00")
    emit s
End Sub

Private Function readText(ByVal path As String) As String
    Dim ff As Integer, b() As Byte, n As Long
    ff = FreeFile
    Open path For Binary Access Read As #ff
    n = LOF(ff)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #ff, 1, b
        readText = utf8_decode(b, 0, n)
    End If
    Close #ff
End Function

Private Sub writeText(ByVal path As String, ByVal s As String)
    Dim ff As Integer, b() As Byte, n As Long
    n = utf8_encode(s, b)
    If Len(Dir$(path)) > 0 Then Kill path
    ff = FreeFile
    Open path For Binary Access Write As #ff
    If n > 0 Then Put #ff, 1, b
    Close #ff
End Sub

Sub Main()
    Dim args() As String, lines() As String, i As Long, outText As String
    QueryPerformanceFrequency freq
    ReDim outLines(0 To 63)
    args = Split(Command$, " ")
    lines = Split(readText(Replace(args(0), """", "")), vbLf)
    For i = 0 To UBound(lines)
        lines(i) = Replace(lines(i), vbCr, "")
        If Len(lines(i)) > 0 Then runCommand lines(i)
    Next
    For i = 0 To outCount - 1
        outText = outText & outLines(i) & vbLf
    Next
    writeText Replace(args(1), """", ""), outText
    ExitProcess 0
End Sub
