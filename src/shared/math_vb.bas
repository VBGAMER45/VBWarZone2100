Attribute VB_Name = "math_vb"
' C's integer arithmetic for the port (no C++ counterpart; compiled into every
' project). A uint32_t is a Long holding the same 32 bits; these helpers do
' what C does with such values: unsigned conversions, compares and shifts.
' Values wider than 32 bits are Doubles (exact below 2^53) or Decimals
' (CDec, exact to 96 bits).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Public Const TWO_32 As Double = 4294967296#
Public Const TWO_53 As Double = 9007199254740992#

' (VB) BIT_'s table of 1 << n (a power of 2 through 2 ^ n is a floating point
' power and a conversion, the slowest part of every shift)
Private bitTable(0 To 31) As Long
Private bitTableReady As Boolean

' The unsigned value of a uint32_t bit pattern.
Public Function U32(ByVal x As Long) As Double
    If x < 0 Then U32 = x + TWO_32 Else U32 = x
End Function

' The uint32_t bit pattern of an integral value, modulo 2^32 (C's conversion to uint32_t / int32_t).
Public Function FromU32(ByVal x As Double) As Long
    x = x - Int(x / TWO_32) * TWO_32
    If x >= 2147483648# Then FromU32 = x - TWO_32 Else FromU32 = x
End Function

' C: a * b for uint32_t / int32_t, modulo 2^32 (the 16-bit halves multiplied in Double, exactly).
Public Function MUL32(ByVal a As Long, ByVal b As Long) As Long
    Dim al As Double, ah As Double, bl As Double, bh As Double, m As Double, r As Double
    al = a And &HFFFF&
    ah = (a And &H7FFF0000) / 65536#
    If a < 0 Then ah = ah + 32768#
    bl = b And &HFFFF&
    bh = (b And &H7FFF0000) / 65536#
    If b < 0 Then bh = bh + 32768#
    m = ah * bl + al * bh                               ' < 2^33
    m = m - Int(m / 65536#) * 65536#                    ' the part that stays below bit 32
    r = al * bl + m * 65536#                            ' < 2^49
    MUL32 = FromU32(r)
End Function

' Decimal version of FromU32, for values beyond 2^53.
Public Function FromU32Dec(ByVal x As Variant) As Long
    Dim m As Variant
    m = CDec(x) - Int(CDec(x) / CDec(TWO_32)) * CDec(TWO_32)
    If m >= CDec(2147483648#) Then FromU32Dec = CDbl(m - CDec(TWO_32)) Else FromU32Dec = CDbl(m)
End Function

' C: (a < b) for uint32_t
Public Function ULT(ByVal a As Long, ByVal b As Long) As Boolean
    ULT = (a Xor &H80000000) < (b Xor &H80000000)
End Function

' C: x << n for uint32_t / int32_t (bits shifted out are lost), 0 <= n <= 31
Public Function SHL(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then
        SHL = x
    ElseIf n >= 32 Then
        SHL = 0
    Else
        x = x And (BIT_(32 - n) - 1)                    ' the bits that survive, below bit 32 - n
        If (x And BIT_(31 - n)) <> 0 Then
            SHL = ((x And (BIT_(31 - n) - 1)) * BIT_(n)) Or &H80000000
        Else
            SHL = x * BIT_(n)
        End If
    End If
End Function

' C: x >> n for uint32_t (logical), 0 <= n <= 31
Public Function SHRU(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then
        SHRU = x
    ElseIf x >= 0 Then
        SHRU = x \ BIT_(n)
    Else
        SHRU = ((x And &H7FFFFFFF) \ BIT_(n)) Or BIT_(31 - n)
    End If
End Function

' C: x >> n for int32_t (arithmetic: rounds toward minus infinity), 0 <= n <= 31
Public Function SAR(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then
        SAR = x
    Else
        SAR = Int(x / BIT_(n))
    End If
End Function

Private Sub InitBitTable()
    Dim i As Long
    bitTable(0) = 1
    For i = 1 To 30
        bitTable(i) = bitTable(i - 1) * 2
    Next
    bitTable(31) = &H80000000                   ' (2 ^ 31 overflows a Long)
    bitTableReady = True
End Sub

' C: 1 << n, 0 <= n <= 31
Public Function BIT_(ByVal n As Long) As Long
    If Not bitTableReady Then InitBitTable
    BIT_ = bitTable(n)
End Function

' C: (int) of a floating point value: truncates; out of range (and NaN) gives 0x80000000, as on x86.
Public Function ToInt(ByVal x As Double) As Long
    If x < 2147483648# And x > -2147483649# Then
        ToInt = Fix(x)
    Else
        ToInt = &H80000000
    End If
End Function

' C: (int8_t) of an int32_t: the low 8 bits, sign-extended.
Public Function ToInt8(ByVal x As Long) As Long
    x = x And &HFF&
    If x >= 128 Then ToInt8 = x - 256 Else ToInt8 = x
End Function

' C: (int16_t) of an int32_t: the low 16 bits, sign-extended.
Public Function ToInt16(ByVal x As Long) As Long
    x = x And &HFFFF&
    If x >= 32768 Then ToInt16 = x - 65536 Else ToInt16 = x
End Function

' A Double from its IEEE 754 bit pattern (hi and lo 32-bit halves).
Public Function DoubleFromBits(ByVal hi As Long, ByVal lo As Long) As Double
    Dim d As Double
    CopyMemory ByVal VarPtr(d), lo, 4
    CopyMemory ByVal VarPtr(d) + 4, hi, 4
    DoubleFromBits = d
End Function
