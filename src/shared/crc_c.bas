Attribute VB_Name = "crc_c"
' lib/framework/crc.cpp (compiled into every project: no state but its
' tables). Ported: crcTable, crcSumU16, crcSumI16, struct Sha256's methods
' (Sha256_<method>(h, ...)), wz::crc_init and wz::crc_update (wz_crc_init,
' wz_crc_update), and crcSumVector2i's step for one int32_t (crcSumI32: the
' exe loops it over its Vector2is; this module has no Vector2i). Missing:
' sha256Sum, EcKey and the session keys (libsodium), base64.
' crcTable is generated with the formula in crc.cpp's comment and checked
' against the table's last entry. wz::crc_update is the reflected CRC-32
' (polynomial 0xEDB88320) without the final XOR: crc_table[0] is generated
' and checked the same way; the other seven tables only let the C++ take 8
' bytes at a time, so the byte-at-a-time loop gives the same result.
Option Explicit

' crc.h: struct Sha256 (Bytes = 32)
Public Type Sha256
    bytes(0 To 31) As Byte
End Type

Private crcTable(0 To 255) As Long
Private crcReady As Boolean

Private crc32Table(0 To 255) As Long     ' crc_table[0]
Private crc32Ready As Boolean

' crcTable[1] = 0x04C11DB7;
' crcTable[i] = crcTable[i>>1]<<1 ^ ((crcTable[i>>1]>>31 ^ (i & 0x01))*crcTable[1]);
Private Sub InitCrcTable()
    Dim i As Long, t As Long
    crcTable(0) = 0
    crcTable(1) = &H4C11DB7
    For i = 2 To 255
        t = crcTable(i \ 2)
        If (SHRU(t, 31) Xor (i And 1)) <> 0 Then
            crcTable(i) = SHL(t, 1) Xor crcTable(1)
        Else
            crcTable(i) = SHL(t, 1)
        End If
    Next
    crcReady = True
    If crcTable(255) <> &HB1F740B4 Then
        debug_ 0, LOG_ERROR, "InitCrcTable", "(VB) the generated crcTable is wrong: crcTable[255] = 0x%08X", crcTable(255)
    End If
End Sub

' One step: crc = crc << 8 ^ crcTable[crc>>24 ^ b]
Private Function CrcStep(ByVal crc As Long, ByVal b As Long) As Long
    CrcStep = SHL(crc, 8) Xor crcTable(SHRU(crc, 24) Xor (b And &HFF&))
End Function

' data(first .. first + dataLen - 1): uint16_t values in Longs
Public Function crcSumU16(ByVal crc As Long, data() As Long, ByVal first As Long, ByVal dataLen As Long) As Long
    Dim i As Long
    If Not crcReady Then InitCrcTable
    For i = first To first + dataLen - 1
        crc = CrcStep(crc, (data(i) And &HFF00&) \ &H100&)
        crc = CrcStep(crc, data(i))
    Next
    crcSumU16 = crc
End Function

' (VB) crcSumVector2i's four steps for one coordinate: (uint8_t)(v >> 24), (v >> 16), (v >> 8), v
Public Function crcSumI32(ByVal crc As Long, ByVal v As Long) As Long
    If Not crcReady Then InitCrcTable
    crc = CrcStep(crc, SHRU(v, 24))
    crc = CrcStep(crc, SHRU(v, 16))
    crc = CrcStep(crc, SHRU(v, 8))
    crcSumI32 = CrcStep(crc, v)
End Function

' data(first .. first + dataLen - 1): int16_t values in Longs
Public Function crcSumI16(ByVal crc As Long, data() As Long, ByVal first As Long, ByVal dataLen As Long) As Long
    Dim i As Long
    If Not crcReady Then InitCrcTable
    For i = first To first + dataLen - 1
        crc = CrcStep(crc, SAR(data(i), 8))
        crc = CrcStep(crc, data(i))
    Next
    crcSumI16 = crc
End Function

' crc_table[0]: the reflected CRC-32 table
Private Sub InitCrc32Table()
    Dim i As Long, k As Long, c As Long
    For i = 0 To 255
        c = i
        For k = 1 To 8
            If (c And 1) <> 0 Then c = SHRU(c, 1) Xor &HEDB88320 Else c = SHRU(c, 1)
        Next
        crc32Table(i) = c
    Next
    crc32Ready = True
    If crc32Table(1) <> &H77073096 Or crc32Table(255) <> &H2D02EF8D Then
        debug_ 0, LOG_ERROR, "InitCrc32Table", "(VB) the generated crc_table[0] is wrong: crc_table[0][255] = 0x%08X", crc32Table(255)
    End If
End Sub

' C: uint32_t wz::crc_init()
Public Function wz_crc_init() As Long
    wz_crc_init = &HFFFFFFFF
End Function

' C: uint32_t wz::crc_update(uint32_t crc, const void *data, size_t data_len): data(first .. first + data_len - 1)
Public Function wz_crc_update(ByVal crc As Long, data() As Byte, ByVal first As Long, ByVal data_len As Long) As Long
    Dim i As Long, tbl_idx As Long
    If Not crc32Ready Then InitCrc32Table
    For i = first To first + data_len - 1
        tbl_idx = (crc Xor data(i)) And &HFF&
        ' (VB) crc >> 8 (unsigned) written out: the inner loop of every CRC
        If crc >= 0 Then
            crc = crc32Table(tbl_idx) Xor (crc \ &H100&)
        Else
            crc = crc32Table(tbl_idx) Xor (((crc And &H7FFFFFFF) \ &H100&) Or &H800000)
        End If
    Next
    wz_crc_update = crc
End Function

' (VB) wz::crc_update over the 4 bytes of a uint32_t in big-endian order (wz_htonl(value)), the steps
' written out: what sync_debug's int lists CRC, 35 values a droid each tick
Public Function wz_crc_update_be32(ByVal crc As Long, ByVal value As Long) As Long
    Dim b As Long, low31 As Long
    If Not crc32Ready Then InitCrc32Table
    low31 = value And &H7FFFFFFF
    b = low31 \ &H1000000                         ' value >> 24
    If value < 0 Then b = b Or &H80&
    If crc >= 0 Then
        crc = crc32Table((crc Xor b) And &HFF&) Xor (crc \ &H100&)
    Else
        crc = crc32Table((crc Xor b) And &HFF&) Xor (((crc And &H7FFFFFFF) \ &H100&) Or &H800000)
    End If
    b = (low31 \ &H10000) And &HFF&               ' value >> 16
    If crc >= 0 Then
        crc = crc32Table((crc Xor b) And &HFF&) Xor (crc \ &H100&)
    Else
        crc = crc32Table((crc Xor b) And &HFF&) Xor (((crc And &H7FFFFFFF) \ &H100&) Or &H800000)
    End If
    b = (low31 \ &H100&) And &HFF&                 ' value >> 8
    If crc >= 0 Then
        crc = crc32Table((crc Xor b) And &HFF&) Xor (crc \ &H100&)
    Else
        crc = crc32Table((crc Xor b) And &HFF&) Xor (((crc And &H7FFFFFFF) \ &H100&) Or &H800000)
    End If
    b = value And &HFF&
    If crc >= 0 Then
        crc = crc32Table((crc Xor b) And &HFF&) Xor (crc \ &H100&)
    Else
        crc = crc32Table((crc Xor b) And &HFF&) Xor (((crc And &H7FFFFFFF) \ &H100&) Or &H800000)
    End If
    wz_crc_update_be32 = crc
End Function

' C: bool Sha256::operator ==(Sha256 const &b) const
Public Function Sha256_equals(h As Sha256, b As Sha256) As Boolean
    Dim n As Long
    For n = 0 To 31
        If h.bytes(n) <> b.bytes(n) Then Exit Function
    Next
    Sha256_equals = True
End Function

Public Function Sha256_isZero(h As Sha256) As Boolean
    Dim n As Long
    For n = 0 To 31
        If h.bytes(n) <> 0 Then Exit Function
    Next
    Sha256_isZero = True
End Function

Public Sub Sha256_setZero(h As Sha256)
    Dim n As Long
    For n = 0 To 31
        h.bytes(n) = 0
    Next
End Sub

' bytesToHex: lower case
Public Function Sha256_toString(h As Sha256) As String
    Const hexDigits As String = "0123456789abcdef"
    Dim str_ As String, n As Long
    str_ = Space$(64)
    For n = 0 To 31
        Mid$(str_, n * 2 + 1, 1) = Mid$(hexDigits, h.bytes(n) \ 16 + 1, 1)
        Mid$(str_, n * 2 + 2, 1) = Mid$(hexDigits, (h.bytes(n) And 15) + 1, 1)
    Next
    Sha256_toString = str_
End Function

Public Sub Sha256_fromString(h As Sha256, ByVal s As String)
    Dim nChars As Long, n As Long, hv As Long, c As Long
    Sha256_setZero h
    nChars = 64
    If Len(s) < nChars Then nChars = Len(s)
    For n = 0 To nChars - 1
        c = AscW(Mid$(s, n + 1, 1)) And &HFFFF&
        If c >= 48 And c <= 57 Then             ' '0'..'9'
            hv = c - 48
        ElseIf c >= 97 And c <= 102 Then        ' 'a'..'f'
            hv = c - 97 + 10
        ElseIf c >= 65 And c <= 70 Then         ' 'A'..'F'
            hv = c - 65 + 10
        Else
            Exit For
        End If
        If n Mod 2 <> 0 Then
            h.bytes(n \ 2) = h.bytes(n \ 2) Or hv
        Else
            h.bytes(n \ 2) = h.bytes(n \ 2) Or (hv * 16)
        End If
    Next
End Sub
