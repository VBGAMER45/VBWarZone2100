Attribute VB_Name = "crc_sodium_c"
' lib/framework/crc.cpp, crc.h: the parts over libsodium (src/shared/sodium: compiled into the exe and WzNet,
' which both use EcKey and SessionKeys; crc_c holds the rest of crc.cpp and is in every project): sha256Sum,
' class EcKey whole, class SessionKeys, base64Encode, base64Decode, b64Tob64UrlSafe, b64UrlSafeTob64,
' genSecRandomBytes. libsodium's functions are sodium_*_c's.
' EcKey is a UDT (C++ copies an EcKey as a value whose key data is shared and never changed: a VB object
' would alias), its methods EcKey_<method>(k, ...); vKey's EC_KEY is privateKey() and publicKey() (vKey !=
' nullptr: hasKey). SessionKeys is a UDT too; its constructor is SessionKeys_init, which returns False where the
' C++ throws std::invalid_argument (the exception's what() is SessionKeys_lastError). An EcKey or a
' SessionKeys crosses between the exe and WzNet as bytes (EcKey_toBytes / EcKey_fromBytes; the two keys).
' A std::vector<unsigned char> result is an array and its size: the arrays are ReDimmed to the size (left
' alone when it is 0).
Option Explicit

' crc.h: EcKey::Privacy
Public Const EcKey_Public As Long = 0
Public Const EcKey_Private As Long = 1

' crc.h: class EcKey (vKey: EC_KEY, or nullptr)
Public Type EcKey
    hasKey As Boolean                   ' vKey != nullptr
    publicKey() As Byte                 ' EC_KEY's publicKey: publicKey(0 .. publicKeySize - 1)
    publicKeySize As Long
    privateKey() As Byte                ' EC_KEY's privateKey (libsodium's ed25519 secret key, seed || public key)
    privateKeySize As Long
End Type

' crc.h: class SessionKeys
Public Type SessionKeys
    receiveKey() As Byte                ' std::vector<unsigned char> receiveKey
    receiveKeySize As Long
    sendKey() As Byte                   ' std::vector<unsigned char> sendKey
    sendKeySize As Long
    buffer() As Byte                    ' to avoid some repeated allocations
    bufferSize As Long
End Type

Private Const crypto_sign_ed25519_BYTES As Long = 64
Private Const crypto_sign_ed25519_PUBLICKEYBYTES As Long = 32
Private Const crypto_sign_ed25519_SECRETKEYBYTES As Long = 64
Private Const crypto_scalarmult_curve25519_BYTES As Long = 32
Private Const crypto_kx_SESSIONKEYBYTES As Long = 32
Private Const crypto_aead_xchacha20poly1305_ietf_NPUBBYTES As Long = 24
Private Const crypto_aead_xchacha20poly1305_ietf_ABYTES As Long = 16
Private Const SessionKeys_NonceSize As Long = 24          ' static constexpr size_t NonceSize = 24
Private Const sessionMessageBlockSize As Long = 64        ' constexpr size_t sessionMessageBlockSize = 64

Private sessionKeysError As String                         ' (VB) the last SessionKeys_init's std::invalid_argument

'--------------------------------------------------------------------------------------------------------
' SHA256

' Sha256 sha256Sum(void const *data, size_t dataLen): data(ofs .. ofs + dataLen - 1)
Public Function sha256Sum(data() As Byte, ByVal ofs As Long, ByVal dataLen As Long) As Sha256
    Dim ret As Sha256
    crypto_hash_sha256 ret.bytes, 0, data, ofs, dataLen
    sha256Sum = ret
End Function

' static inline std::string bytesToHex(const uint8_t* bytes, size_t numBytes)
Private Function bytesToHex(bytes() As Byte, ByVal numBytes As Long) As String
    Const hexDigits As String = "0123456789abcdef"
    Dim str_ As String, n As Long
    str_ = Space$(numBytes * 2)
    For n = 0 To numBytes - 1
        Mid$(str_, n * 2 + 1, 1) = Mid$(hexDigits, bytes(n) \ 16 + 1, 1)
        Mid$(str_, n * 2 + 2, 1) = Mid$(hexDigits, (bytes(n) And 15) + 1, 1)
    Next
    bytesToHex = str_
End Function

'--------------------------------------------------------------------------------------------------------
' ECDSA (ed25519)

Private Sub copyBytes(dst() As Byte, src() As Byte, ByVal n As Long)
    Dim i As Long
    If n <= 0 Then
        Erase dst
        Exit Sub
    End If
    ReDim dst(0 To n - 1)
    For i = 0 To n - 1
        dst(i) = src(i)
    Next
End Sub

' void EcKey::clear()
Public Sub EcKey_clear(k As EcKey)
    k.hasKey = False
    Erase k.publicKey
    k.publicKeySize = 0
    Erase k.privateKey
    k.privateKeySize = 0
End Sub

' bool EcKey::empty() const
Public Function EcKey_empty(k As EcKey) As Boolean
    EcKey_empty = Not k.hasKey
End Function

' bool EcKey::hasPrivate() const
Public Function EcKey_hasPrivate(k As EcKey) As Boolean
    EcKey_hasPrivate = k.hasKey And k.privateKeySize > 0
End Function

' Sig EcKey::sign(void const *data, size_t dataLen) const: the signature into sig(0 .. returned size - 1) (0: an
' empty Sig)
Public Function EcKey_sign(k As EcKey, data() As Byte, ByVal ofs As Long, ByVal dataLen As Long, sig() As Byte) As Long
    Dim signature() As Byte, privateKey() As Byte, signSuccess As Long
    If Not k.hasKey Or k.privateKeySize = 0 Then
        debug_ 258, LOG_ERROR, "sign", "No key"
        EcKey_sign = 0
        Exit Function
    End If

    ReDim signature(0 To crypto_sign_ed25519_BYTES - 1)

    copyBytes privateKey, k.privateKey, k.privateKeySize

    signSuccess = crypto_sign_ed25519_detached(signature, 0, data, ofs, dataLen, privateKey, 0)
    If signSuccess <> 0 Then
        debug_ 278, LOG_ERROR, "sign", "Creating a signature failed"
        EcKey_sign = 0
        Exit Function
    End If

    copyBytes sig, signature, crypto_sign_ed25519_BYTES
    EcKey_sign = crypto_sign_ed25519_BYTES
End Function

' bool EcKey::verify(Sig const &sig, void const *data, size_t dataLen) const: sig(0 .. sigLen - 1)
Public Function EcKey_verify(k As EcKey, sig() As Byte, ByVal sigLen As Long, data() As Byte, ByVal ofs As Long, ByVal dataLen As Long) As Boolean
    Dim verifyResult As Long
    If Not k.hasKey Or k.publicKeySize = 0 Then
        debug_ 289, LOG_ERROR, "verify", "No key"
        EcKey_verify = False
        Exit Function
    End If

    If sigLen <> crypto_sign_ed25519_BYTES Then
        debug_ 295, LOG_INFO, "verify", "Signature is the wrong size (received: %zu, expecting: %zu)", sigLen, crypto_sign_ed25519_BYTES
        EcKey_verify = False
        Exit Function
    End If

    verifyResult = crypto_sign_ed25519_verify_detached(sig, 0, data, ofs, dataLen, k.publicKey, 0)
    If verifyResult <> 0 Then
        EcKey_verify = False
        Exit Function
    End If

    EcKey_verify = True
End Function

' Key EcKey::toBytes(Privacy privacy) const: the key into key(0 .. returned size - 1) (0: an empty Key)
Public Function EcKey_toBytes(k As EcKey, ByVal privacy As Long, key() As Byte) As Long
    If EcKey_empty(k) Then
        debug_ 321, LOG_WZ, "toBytes", "No key"
        EcKey_toBytes = 0
        Exit Function
    End If

    If privacy = EcKey_Public Then
        copyBytes key, k.publicKey, k.publicKeySize
        EcKey_toBytes = k.publicKeySize
    ElseIf privacy = EcKey_Private Then
        If k.privateKeySize <> crypto_sign_ed25519_SECRETKEYBYTES Then
            debug_ 335, LOG_ERROR, "toBytes", "Failed to create external representation of private key"
            EcKey_toBytes = 0
            Exit Function
        End If

        copyBytes key, k.privateKey, k.privateKeySize
        EcKey_toBytes = k.privateKeySize
    Else
        debug_ 343, LOG_FATAL, "toBytes", "Unsupported privacy level"
        EcKey_toBytes = 0
    End If
End Function

' toBytes(EcKey::Public) (the exe's earlier name for it)
Public Function EcKey_toBytes_public(k As EcKey, key() As Byte) As Long
    EcKey_toBytes_public = EcKey_toBytes(k, EcKey_Public, key)
End Function

' bool EcKey::fromBytes(EcKey::Key const &key, EcKey::Privacy privacy): key(0 .. keySize - 1)
Public Function EcKey_fromBytes(k As EcKey, key() As Byte, ByVal keySize As Long, ByVal privacy As Long) As Boolean
    Dim publicKey() As Byte
    EcKey_clear k

    If privacy = EcKey_Public Then
        If keySize <> crypto_sign_ed25519_PUBLICKEYBYTES Then
            debug_ 356, LOG_ERROR, "fromBytes", "Invalid public key"
            EcKey_fromBytes = False
            Exit Function
        End If

        If crypto_core_ed25519_is_valid_point(key, 0) = 0 Then
            debug_ 362, LOG_ERROR, "fromBytes", "Invalid public key"
            EcKey_fromBytes = False
            Exit Function
        End If

        copyBytes k.publicKey, key, keySize    ' std::make_shared<EC_KEY>(std::vector<unsigned char>(0), key)
        k.publicKeySize = keySize
        k.hasKey = True
    ElseIf privacy = EcKey_Private Then
        ' Verify private key is expected length
        If keySize <> crypto_sign_ed25519_SECRETKEYBYTES Then
            ' Invalid input
            debug_ 374, LOG_ERROR, "fromBytes", "Invalid private key length"
            EcKey_fromBytes = False
            Exit Function
        End If

        ' Compute public key from private key
        ReDim publicKey(0 To crypto_sign_ed25519_PUBLICKEYBYTES - 1)
        If crypto_sign_ed25519_sk_to_pk(publicKey, 0, key, 0) <> 0 Then
            ' Failed to compute public key from private key
            debug_ 383, LOG_ERROR, "fromBytes", "Invalid private key"
            EcKey_fromBytes = False
            Exit Function
        End If

        If crypto_core_ed25519_is_valid_point(publicKey, 0) = 0 Then
            debug_ 389, LOG_ERROR, "fromBytes", "Invalid keypair"
            EcKey_fromBytes = False
            Exit Function
        End If

        copyBytes k.privateKey, key, keySize    ' std::make_shared<EC_KEY>(key, publicKey)
        k.privateKeySize = keySize
        k.publicKey = publicKey
        k.publicKeySize = crypto_sign_ed25519_PUBLICKEYBYTES
        k.hasKey = True
    Else
        debug_ 397, LOG_FATAL, "fromBytes", "Unsupported privacy level"
        EcKey_fromBytes = False
        Exit Function
    End If

    EcKey_fromBytes = True
End Function

' fromBytes(key, EcKey::Public) (the exe's earlier name for it)
Public Function EcKey_fromBytes_public(k As EcKey, key() As Byte, ByVal keySize As Long) As Boolean
    EcKey_fromBytes_public = EcKey_fromBytes(k, key, keySize, EcKey_Public)
End Function

' static EcKey EcKey::generate(): into k (an empty key when it fails)
Public Sub EcKey_generate(k As EcKey)
    Dim publicKey() As Byte, privateKey() As Byte, genResult As Long
    EcKey_clear k
    ' EC_KEY::createAndReserveForCurve()
    ReDim privateKey(0 To crypto_sign_ed25519_SECRETKEYBYTES - 1)
    ReDim publicKey(0 To crypto_sign_ed25519_PUBLICKEYBYTES - 1)

    genResult = crypto_sign_ed25519_keypair(publicKey, 0, privateKey, 0)
    If genResult <> 0 Then
        ' crypto_sign_ed25519_keypair failed
        debug_ 412, LOG_ERROR, "generate", "Failed to generate key pair"
        Exit Sub
    End If

    k.privateKey = privateKey
    k.privateKeySize = crypto_sign_ed25519_SECRETKEYBYTES
    k.publicKey = publicKey
    k.publicKeySize = crypto_sign_ed25519_PUBLICKEYBYTES
    k.hasKey = True
End Sub

' std::string EcKey::publicHashString(size_t truncateToLength = 0) const
Public Function EcKey_publicHashString(k As EcKey, Optional ByVal truncateToLength As Long = 0) As String
    Dim bytes() As Byte, n As Long, shaStr As String
    n = EcKey_toBytes(k, EcKey_Public, bytes)
    If n = 0 Then
        EcKey_publicHashString = ""
        Exit Function
    End If
    shaStr = Sha256_toString(sha256Sum(bytes, 0, n))
    If truncateToLength > 0 And truncateToLength < Len(shaStr) Then
        EcKey_publicHashString = Left$(shaStr, truncateToLength)
        Exit Function
    End If
    EcKey_publicHashString = shaStr
End Function

' std::string EcKey::publicKeyHexString(size_t truncateToLength = 0) const (4.7.0 does not truncate)
Public Function EcKey_publicKeyHexString(k As EcKey, Optional ByVal truncateToLength As Long = 0) As String
    Dim bytes() As Byte, n As Long
    n = EcKey_toBytes(k, EcKey_Public, bytes)
    If n = 0 Then
        EcKey_publicKeyHexString = ""
        Exit Function
    End If
    EcKey_publicKeyHexString = bytesToHex(bytes, n)
End Function

'--------------------------------------------------------------------------------------------------------
' Base64

' std::string base64Encode(std::vector<uint8_t> const &bytes): bytes(0 .. size - 1)
Public Function base64Encode(bytes() As Byte, ByVal size As Long) As String
    Const alphabet As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim str_ As String, n As Long, rem_ As Long, block As Long, i As Long, b1 As Long, b2 As Long
    str_ = String$(((size + 2) \ 3) * 4, vbNullChar)
    n = 0
    Do While n * 3 < size
        rem_ = size - n * 3
        If rem_ > 1 Then b1 = bytes(1 + n * 3) Else b1 = 0
        If rem_ > 2 Then b2 = bytes(2 + n * 3) Else b2 = 0
        block = CLng(bytes(0 + n * 3)) * 65536 Or b1 * 256 Or b2
        For i = 0 To 3
            Mid$(str_, i + n * 4 + 1, 1) = Mid$(alphabet, ((block \ BIT_(6 * (3 - i))) And &H3F) + 1, 1)
        Next
        If rem_ <= 2 Then
            Mid$(str_, 3 + n * 4 + 1, 1) = "="
            If rem_ <= 1 Then
                Mid$(str_, 2 + n * 4 + 1, 1) = "="
            End If
        End If
        n = n + 1
    Loop
    base64Encode = str_
End Function

' std::vector<uint8_t> base64Decode(std::string const &str): the bytes into bytes(0 .. returned size - 1). (VB) str
' is the std::string's bytes: its characters as UTF-8 (base64 text is ASCII)
Public Function base64Decode(ByVal str_ As String, bytes() As Byte) As Long
    Dim s() As Byte, sLen As Long, size As Long, n As Long, i As Long, block As Long, ch As Long, val_ As Long
    sLen = utf8_encode(str_, s)
    size = (sLen \ 4) * 3
    If size > 0 Then ReDim bytes(0 To size - 1)
    For n = 0 To (sLen \ 4) - 1
        block = 0
        For i = 0 To 3
            ch = s(i + n * 4)
            If ch >= 65 And ch <= 90 Then           ' 'A'..'Z'
                val_ = ch - 65 + 0
            ElseIf ch >= 97 And ch <= 122 Then      ' 'a'..'z'
                val_ = ch - 97 + 26
            ElseIf ch >= 48 And ch <= 57 Then       ' '0'..'9'
                val_ = ch - 48 + 52
            ElseIf ch = 43 Then                     ' '+'
                val_ = 62
            ElseIf ch = 47 Then                     ' '/'
                val_ = 63
            Else
                val_ = 0
            End If
            block = block Or val_ * BIT_(6 * (3 - i))
        Next
        bytes(0 + n * 3) = (block \ 65536) And &HFF&
        bytes(1 + n * 3) = (block \ 256) And &HFF&
        bytes(2 + n * 3) = block And &HFF&
    Next
    If sLen >= 4 Then
        If s((sLen \ 4) * 4 - 2) = 61 Then            ' '='
            size = size - 2
        ElseIf s((sLen \ 4) * 4 - 1) = 61 Then
            size = size - 1
        End If
    End If
    ' bytes.resize(...): the array is the decoded bytes, as the C++ vector (a padded input had two or one more)
    If size > 0 And size <> (sLen \ 4) * 3 Then ReDim Preserve bytes(0 To size - 1)
    base64Decode = size
End Function

' std::string b64Tob64UrlSafe(const std::string& inputb64)
Public Function b64Tob64UrlSafe(ByVal inputb64 As String) As String
    Dim b64urlsafe As String
    b64urlsafe = inputb64
    b64urlsafe = Replace(b64urlsafe, "+", "-")
    b64urlsafe = Replace(b64urlsafe, "/", "_")
    b64Tob64UrlSafe = b64urlsafe
End Function

' std::string b64UrlSafeTob64(const std::string& inputb64urlsafe)
Public Function b64UrlSafeTob64(ByVal inputb64urlsafe As String) As String
    Dim b64 As String
    b64 = inputb64urlsafe
    b64 = Replace(b64, "-", "+")
    b64 = Replace(b64, "_", "/")
    b64UrlSafeTob64 = b64
End Function

'--------------------------------------------------------------------------------------------------------
' Generating Random Data

' void genSecRandomBytes(void * const buf, const size_t size): buf(ofs .. ofs + size - 1) (the
' std::vector<uint8_t> genSecRandomBytes(size_t numBytes) form: ReDim the array, then this)
Public Sub genSecRandomBytes(buf() As Byte, ByVal ofs As Long, ByVal size As Long)
    randombytes_buf buf, ofs, size
End Sub

'--------------------------------------------------------------------------------------------------------
' SessionKeys

' (VB) what() of the std::invalid_argument the last SessionKeys_init threw ("" when it did not)
Public Function SessionKeys_lastError() As String
    SessionKeys_lastError = sessionKeysError
End Function

' SessionKeys::SessionKeys(EcKey const &me, uint32_t me_playerIdx, EcKey const &other, uint32_t other_playerIdx):
' False where it throws std::invalid_argument
Public Function SessionKeys_init(sk As SessionKeys, me_ As EcKey, ByVal me_playerIdx As Long, other As EcKey, ByVal other_playerIdx As Long) As Boolean
    Dim me_publicKeyStr As String, other_publicKeyStr As String, bIsClient As Boolean
    Dim ed25519_pk() As Byte, ed25519_sk() As Byte, me_pk() As Byte, me_sk() As Byte
    Dim other_ed25519_pk() As Byte, other_pk() As Byte, n As Long, i As Long
    sessionKeysError = ""
    Erase sk.receiveKey
    sk.receiveKeySize = 0
    Erase sk.sendKey
    sk.sendKeySize = 0

    If Not EcKey_hasPrivate(me_) Then
        sessionKeysError = "Invalid ""me"" key"         ' throw std::invalid_argument
        Exit Function
    End If

    ' The libsodium API expects a deterministic "client" and "server"
    ' To do so, sort the EcKeys by public key hex string
    ' If they are somehow both equal (probably someone testing locally with the same config dir), use the player index instead
    me_publicKeyStr = EcKey_publicKeyHexString(me_)
    other_publicKeyStr = EcKey_publicKeyHexString(other)
    If me_publicKeyStr <> other_publicKeyStr Then
        bIsClient = (StrComp(me_publicKeyStr, other_publicKeyStr, vbBinaryCompare) < 0)
    Else
        bIsClient = ULT(me_playerIdx, other_playerIdx)
    End If

    EcKey_toBytes me_, EcKey_Public, ed25519_pk
    EcKey_toBytes me_, EcKey_Private, ed25519_sk
    ReDim me_pk(0 To crypto_scalarmult_curve25519_BYTES - 1)
    ReDim me_sk(0 To crypto_scalarmult_curve25519_BYTES - 1)
    If crypto_sign_ed25519_pk_to_curve25519(me_pk, 0, ed25519_pk, 0) <> 0 Then
        sessionKeysError = "Invalid public key"
        Exit Function
    End If
    If crypto_sign_ed25519_sk_to_curve25519(me_sk, 0, ed25519_sk, 0) <> 0 Then
        sessionKeysError = "Invalid private key"
        Exit Function
    End If
    For i = 0 To UBound(ed25519_sk)             ' sodium_memzero(&ed25519_sk[0], ed25519_sk.size())
        ed25519_sk(i) = 0
    Next
    Erase ed25519_sk

    ' (VB) an empty "other" key: C++ hands libsodium the empty vector's data() (null)
    n = EcKey_toBytes(other, EcKey_Public, other_ed25519_pk)
    ReDim other_pk(0 To crypto_scalarmult_curve25519_BYTES - 1)
    If n <> crypto_sign_ed25519_PUBLICKEYBYTES Then
        sessionKeysError = "Invalid ""other"" public key"
        Exit Function
    End If
    If crypto_sign_ed25519_pk_to_curve25519(other_pk, 0, other_ed25519_pk, 0) <> 0 Then
        sessionKeysError = "Invalid ""other"" public key"
        Exit Function
    End If

    ReDim sk.receiveKey(0 To crypto_kx_SESSIONKEYBYTES - 1)
    ReDim sk.sendKey(0 To crypto_kx_SESSIONKEYBYTES - 1)

    If bIsClient Then
        If crypto_kx_client_session_keys(sk.receiveKey, 0, sk.sendKey, 0, me_pk, 0, me_sk, 0, other_pk, 0) <> 0 Then
            ' Suspicious other public key
            Erase sk.receiveKey
            Erase sk.sendKey
            sessionKeysError = "Invalid ""other"" public key"
            Exit Function
        End If
    Else
        If crypto_kx_server_session_keys(sk.receiveKey, 0, sk.sendKey, 0, me_pk, 0, me_sk, 0, other_pk, 0) <> 0 Then
            ' Suspicious other public key
            Erase sk.receiveKey
            Erase sk.sendKey
            sessionKeysError = "Invalid ""other"" public key"
            Exit Function
        End If
    End If
    sk.receiveKeySize = crypto_kx_SESSIONKEYBYTES
    sk.sendKeySize = crypto_kx_SESSIONKEYBYTES

    For i = 0 To UBound(me_sk)                  ' sodium_memzero(&me_sk[0], me_sk.size())
        me_sk(i) = 0
    Next
    Erase me_sk
    SessionKeys_init = True
End Function

' std::vector<uint8_t> SessionKeys::encryptMessageForOther(void const *data, size_t dataLen): <nonce><ciphertext>
' into out(0 .. returned size - 1) (0: an empty vector)
Public Function SessionKeys_encryptMessageForOther(sk As SessionKeys, data() As Byte, ByVal ofs As Long, ByVal dataLen As Long, out() As Byte) As Long
    Dim bufferUnpaddedLen As Long, bufferPaddedLen As Long, bufferMaxLen As Long, i As Long
    Dim ciphertext() As Byte, ciphertext_len As Long, nonce(0 To SessionKeys_NonceSize - 1) As Byte, noAd() As Byte
    Dim fullResult() As Byte

    bufferUnpaddedLen = dataLen
    bufferPaddedLen = 0
    bufferMaxLen = dataLen + (sessionMessageBlockSize - (dataLen Mod sessionMessageBlockSize))
    ReDim sk.buffer(0 To bufferMaxLen - 1)      ' buffer.clear(); the data copied in; buffer.resize(bufferMaxLen, 0)
    For i = 0 To dataLen - 1
        sk.buffer(i) = data(ofs + i)
    Next
    sk.bufferSize = bufferMaxLen
    ' Round the length of the buffer to a multiple of `sessionMessageBlockSize` by appending padding data
    If sodium_pad(bufferPaddedLen, sk.buffer, dataLen, sessionMessageBlockSize, sk.bufferSize) <> 0 Then
        ASSERT_FAILURE "false", "lib/framework/crc.cpp:631", "encryptMessageForOther", "Not large enough output buffer for padding"
        SessionKeys_encryptMessageForOther = 0
        Exit Function
    End If
    If Not (bufferPaddedLen >= bufferUnpaddedLen) Then ASSERT_FAILURE "bufferPaddedLen >= bufferUnpaddedLen", "lib/framework/crc.cpp:634", "encryptMessageForOther", "Should not be smaller?"
    sk.bufferSize = bufferPaddedLen             ' buffer.resize(bufferPaddedLen)

    ReDim ciphertext(0 To bufferPaddedLen + crypto_aead_xchacha20poly1305_ietf_ABYTES - 1)
    ciphertext_len = 0

    randombytes_buf nonce, 0, SessionKeys_NonceSize

    If crypto_aead_xchacha20poly1305_ietf_encrypt(ciphertext, 0, ciphertext_len, sk.buffer, 0, sk.bufferSize, noAd, 0, 0, nonce, 0, sk.sendKey, 0) <> 0 Then
        ' failed to encrypt
        debug_ 649, LOG_INFO, "encryptMessageForOther", "Encryption failed?"
        SessionKeys_encryptMessageForOther = 0
        Exit Function
    End If

    ' Return <nonce><ciphertext>
    ReDim fullResult(0 To SessionKeys_NonceSize + bufferPaddedLen + crypto_aead_xchacha20poly1305_ietf_ABYTES - 1)
    For i = 0 To SessionKeys_NonceSize - 1
        fullResult(i) = nonce(i)
    Next
    For i = 0 To bufferPaddedLen + crypto_aead_xchacha20poly1305_ietf_ABYTES - 1
        fullResult(SessionKeys_NonceSize + i) = ciphertext(i)
    Next
    out = fullResult
    SessionKeys_encryptMessageForOther = UBound(fullResult) + 1
End Function

' bool SessionKeys::decryptMessageFromOther(void const *data, size_t dataLen, std::vector<uint8_t>& decrypted): the
' message into outputDecrypted(0 .. outLen - 1)
Public Function SessionKeys_decryptMessageFromOther(sk As SessionKeys, data() As Byte, ByVal ofs As Long, ByVal dataLen As Long, outputDecrypted() As Byte, outLen As Long) As Boolean
    Dim ciphertextLen As Long, decryptedLen As Long, bufferPaddedLen As Long, bufferUnpaddedLen As Long
    Dim decrypted() As Byte, noAd() As Byte, i As Long
    outLen = 0
    If Not (dataLen > crypto_aead_xchacha20poly1305_ietf_NPUBBYTES) Then ASSERT_FAILURE "dataLen > crypto_aead_xchacha20poly1305_ietf_NPUBBYTES", "lib/framework/crc.cpp:663", "decryptMessageFromOther", "Invalid dataLen": Exit Function
    ' (the C++'s dataLen < uint32 max: a Long always is)

    ' Split data: <nonce><ciphertext>
    ciphertextLen = dataLen - crypto_aead_xchacha20poly1305_ietf_NPUBBYTES
    If Not (ciphertextLen >= crypto_aead_xchacha20poly1305_ietf_ABYTES) Then ASSERT_FAILURE "ciphertextLen >= crypto_aead_xchacha20poly1305_ietf_ABYTES", "lib/framework/crc.cpp:670", "decryptMessageFromOther", "Invalid dataLen": Exit Function

    ReDim decrypted(0 To ciphertextLen - 1)     ' decrypted.resize(ciphertextLen, 0)
    decryptedLen = 0

    If crypto_aead_xchacha20poly1305_ietf_decrypt(decrypted, 0, decryptedLen, data, ofs + crypto_aead_xchacha20poly1305_ietf_NPUBBYTES, ciphertextLen, noAd, 0, 0, data, ofs, sk.receiveKey, 0) <> 0 Then
        ' Invalid / forged message
        ASSERT_FAILURE "false", "lib/framework/crc.cpp:682", "decryptMessageFromOther", "Invalid encrypted message"
        SessionKeys_decryptMessageFromOther = False
        Exit Function
    End If

    ' Compute the original, unpadded length
    bufferPaddedLen = decryptedLen
    bufferUnpaddedLen = 0
    If sodium_unpad(bufferUnpaddedLen, decrypted, bufferPaddedLen, sessionMessageBlockSize) <> 0 Then
        ASSERT_FAILURE "false", "lib/framework/crc.cpp:691", "decryptMessageFromOther", "Incorrect padding"
        SessionKeys_decryptMessageFromOther = False
        Exit Function
    End If

    If bufferUnpaddedLen > 0 Then               ' decrypted.resize(bufferUnpaddedLen)
        ReDim outputDecrypted(0 To bufferUnpaddedLen - 1)
        For i = 0 To bufferUnpaddedLen - 1
            outputDecrypted(i) = decrypted(i)
        Next
    End If
    outLen = bufferUnpaddedLen
    SessionKeys_decryptMessageFromOther = True
End Function
