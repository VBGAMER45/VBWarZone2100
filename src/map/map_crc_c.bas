Attribute VB_Name = "map_crc_c"
' lib/wzmaplib/src/map_crc.cpp, map_crc.h: the CRC of map objects, in the
' library's own namespace (lib/framework/crc.cpp's code). Ported: crcTable
' (generated from the invariant in its comment and checked against its
' last entry), crcSum, crcSumU32, crcSumWorldPos (one WorldPos, as its
' fields: every call passes dataLen 1). crcSumU16 and crcSumI16 are
' crc_c's (src/shared), the same code.
Option Explicit

Private crcTable(0 To 255) As Long
Private crcReady As Boolean

' Invariant:
' crcTable[0] = 0;
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

' crc = crc << 8 ^ crcTable[crc>>24 ^ (uint8_t)b]
Private Function CrcStep(ByVal crc As Long, ByVal b As Long) As Long
    CrcStep = SHL(crc, 8) Xor crcTable(SHRU(crc, 24) Xor (b And &HFF&))
End Function

' data(first .. first + dataLen - 1)
Public Function crcSum(ByVal crc As Long, data() As Byte, ByVal first As Long, ByVal dataLen As Long) As Long
    Dim i As Long
    If Not crcReady Then InitCrcTable
    For i = first To first + dataLen - 1
        crc = CrcStep(crc, data(i))
    Next
    crcSum = crc
End Function

' data(first .. first + dataLen - 1): uint32_t values in Longs
Public Function crcSumU32(ByVal crc As Long, data() As Long, ByVal first As Long, ByVal dataLen As Long) As Long
    Dim i As Long
    If Not crcReady Then InitCrcTable
    For i = first To first + dataLen - 1
        crc = CrcStep(crc, SHRU(data(i), 24))
        crc = CrcStep(crc, SHRU(data(i), 16))
        crc = CrcStep(crc, SHRU(data(i), 8))
        crc = CrcStep(crc, data(i))
    Next
    crcSumU32 = crc
End Function

' one WorldPos {x, y}
Public Function crcSumWorldPos(ByVal crc As Long, ByVal x As Long, ByVal y As Long) As Long
    If Not crcReady Then InitCrcTable
    crc = CrcStep(crc, SHRU(x, 24))
    crc = CrcStep(crc, SHRU(x, 16))
    crc = CrcStep(crc, SHRU(x, 8))
    crc = CrcStep(crc, x)
    crc = CrcStep(crc, SHRU(y, 24))
    crc = CrcStep(crc, SHRU(y, 16))
    crc = CrcStep(crc, SHRU(y, 8))
    crc = CrcStep(crc, y)
    crcSumWorldPos = crc
End Function
