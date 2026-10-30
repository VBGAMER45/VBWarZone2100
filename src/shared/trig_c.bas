Attribute VB_Name = "trig_c"
' lib/framework/trig.cpp, with trig.h's angleDelta (compiled into every
' project: the simulation calls these on its hot paths). Each project builds
' the tables on first use; trigInitialise (frameInitialise) builds them and
' runs the C++ checks: the tables' CRCs prove the VB arithmetic matches C's.
' uint16_t values are Longs in 0..65535. 64-bit products below 2^53 are
' exact in a Double; i64Sqrt takes a Double below 2^53 or a Decimal above.
Option Explicit

Private trigSinTable(0 To &H4000&) As Long
Private trigAtanTable(0 To &H2000&) As Long
Private trigReady As Boolean

Private Sub BuildTables()
    Dim i As Long, M_PI As Double
    M_PI = DoubleFromBits(&H400921FB, &H54442D18)       ' the double nearest pi, as <math.h>'s M_PI
    For i = 0 To &H4000&
        trigSinTable(i) = (ToInt(&H10000 * Sin(i * (M_PI / &H8000&)) + 0.5) - Abs(i <> 0)) And &HFFFF&    ' -!!i = subtract 1, unless i == 0.
    Next
    For i = 0 To &H2000&
        trigAtanTable(i) = ToInt(&H8000& / M_PI * Atn(CDbl(i) / &H2000&) + 0.5) And &HFFFF&
    Next
    trigReady = True
End Sub

' Initialise the Trig tables
Public Function trigInitialise() As Boolean
    Dim crc As Long, test As Long, lower As Double, upper As Double
    Dim test64 As Variant, lower64 As Variant, upper64 As Variant, r As Long

    ' Generate tables.
    BuildTables

    ' Check tables are correct.
    crc = Not crcSumU16(0, trigSinTable, 0, &H4001&)
    If Not (crc = &H6D3C8DB5) Then ASSERT_FAILURE "crc == 0x6D3C8DB5", "lib/framework/trig.cpp:71", "trigInitialise", "Bad trigSinTable CRC = 0x%08X, sin function is broken.", crc
    crc = Not crcSumU16(0, trigAtanTable, 0, &H2001&)
    If Not (crc = &HD2797F85) Then ASSERT_FAILURE "crc == 0xD2797F85", "lib/framework/trig.cpp:73", "trigInitialise", "Bad trigAtanTable CRC = 0x%08X, atan function is broken.", crc

    ' Test large and small square roots.
    For test = 0 To &HFFFF&
        lower = CDbl(test) * test
        upper = (CDbl(test) + 1) * (test + 1) - 1
        r = iSqrt(FromU32(lower))
        If Not (r = test) Then ASSERT_FAILURE "(uint32_t)iSqrt(lower) == test", "lib/framework/trig.cpp:80", "trigInitialise", "Sanity check failed, iSqrt(%u) gave %u instead of %u!", lower, r, test
        r = iSqrt(FromU32(upper))
        If Not (r = test) Then ASSERT_FAILURE "(uint32_t)iSqrt(upper) == test", "lib/framework/trig.cpp:81", "trigInitialise", "Sanity check failed, iSqrt(%u) gave %u instead of %u!", upper, r, test
    Next
    ' for (uint64_t test = 0xFFFE0000; test != 0x00020000; test = (test + 1) & 0xFFFFFFFF)
    test64 = CDec(4294836224#)
    Do While test64 <> CDec(131072)
        lower64 = test64 * test64
        upper64 = (test64 + 1) * (test64 + 1) - 1
        If upper64 >= CDec(TWO_32) * CDec(TWO_32) Then upper64 = upper64 - CDec(TWO_32) * CDec(TWO_32)   ' uint64_t wraps
        r = i64Sqrt(lower64)
        If Not (r = FromU32Dec(test64)) Then ASSERT_FAILURE "(uint32_t)i64Sqrt(lower) == test", "lib/framework/trig.cpp:87", "trigInitialise", "Sanity check failed, i64Sqrt(%u) gave %u instead of %u!", lower64, r, test64
        r = i64Sqrt(upper64)
        If Not (r = FromU32Dec(test64)) Then ASSERT_FAILURE "(uint32_t)i64Sqrt(upper) == test", "lib/framework/trig.cpp:88", "trigInitialise", "Sanity check failed, i64Sqrt(%u) gave %u instead of %u!", upper64, r, test64
        test64 = test64 + 1
        If test64 >= CDec(TWO_32) Then test64 = test64 - CDec(TWO_32)
    Loop

    trigInitialise = True
End Function

' Returns sin(a*2pi >> 16) << 16, rounded to nearest integer. a: uint16_t (only its low 16 bits count).
Public Function iSin(ByVal a As Long) As Long
    Dim q As Long, r As Long, rvr As Long, v As Long
    If Not trigReady Then BuildTables
    a = a And &HFFFF&
    q = a \ &H4000&
    r = a And &H3FFF&
    If q = 1 Or q = 3 Then rvr = &H4000& - r Else rvr = r            ' reverse = {false, true, false, true}
    v = trigSinTable(rvr)
    If rvr <> 0 Then v = v + 1                                       ' +!!rvr = add 1, unless rvr == 0.
    If q >= 2 Then iSin = -v Else iSin = v                           ' sign = {1, 1, -1, -1}
End Function

' Returns cos(a*2pi >> 16) << 16, rounded to nearest integer.
Public Function iCos(ByVal a As Long) As Long
    Dim q As Long, r As Long, rvr As Long, v As Long
    If Not trigReady Then BuildTables
    a = a And &HFFFF&
    q = a \ &H4000&
    r = a And &H3FFF&
    If q = 0 Or q = 2 Then rvr = &H4000& - r Else rvr = r            ' reverse = {true, false, true, false}
    v = trigSinTable(rvr)
    If rvr <> 0 Then v = v + 1
    If q = 1 Or q = 2 Then iCos = -v Else iCos = v                   ' sign = {1, -1, -1, 1}
End Function

' Returns r*sin(a*2pi >> 16), with up to 16 bits precision.
Public Function iSinR(ByVal a As Long, ByVal r As Long) As Long
    iSinR = Fix(CDbl(r) * iSin(a) / 65536#)
End Function

' Returns r*cos(a*2pi >> 16), with up to 16 bits precision.
Public Function iCosR(ByVal a As Long, ByVal r As Long) As Long
    iCosR = Fix(CDbl(r) * iCos(a) / 65536#)
End Function

' ((int64_t)a << 16) / s, converted to uint16_t
Private Function ScaledAngle(ByVal a As Long, ByVal s As Long) As Long
    ScaledAngle = FromU32(Fix(CDbl(a) * 65536# / s)) And &HFFFF&
End Function

' Returns r*sin(a*2pi/s), with up to 16 bits precision.
Public Function iSinSR(ByVal a As Long, ByVal s As Long, ByVal r As Long) As Long
    iSinSR = Fix(CDbl(r) * iSin(ScaledAngle(a, s)) / 65536#)
End Function

' Returns r*cos(a*2pi/s), with up to 16 bits precision.
Public Function iCosSR(ByVal a As Long, ByVal s As Long, ByVal r As Long) As Long
    iCosSR = Fix(CDbl(r) * iCos(ScaledAngle(a, s)) / 65536#)
End Function

' Returns atan2(s, c)/2pi << 16 as a uint16_t, with a small +-1.5 platform-independent error.
Public Function iAtan2(ByVal s As Long, ByVal c As Long) As Long
    Dim d As Long, j As Double, k As Double

    If s = 0 And c = 0 Then
        iAtan2 = 0       ' All return values equally valid, don't divide by 0.
        Exit Function
    End If
    If Not trigReady Then BuildTables

    ' j and k are uint32_t: the negations wrap as in C (-INT_MIN stays 0x80000000)
    Select Case (-(s < 0) * 2) + -(c < 0)
    Case 0: j = U32(s): k = U32(c): d = &H0
    Case 1: j = U32(FromU32(-CDbl(c))): k = U32(s): d = &H4000&
    Case 2: j = U32(c): k = U32(FromU32(-CDbl(s))): d = &HC000&
    Case 3: j = U32(FromU32(-CDbl(s))): k = U32(FromU32(-CDbl(c))): d = &H8000&
    End Select
    If j < k Then
        iAtan2 = (d + trigAtanTable(Int((j * &H2000& + Int(k / 2)) / k))) And &HFFFF&
    Else
        iAtan2 = (d + &H4000& - trigAtanTable(Int((k * &H2000& + Int(j / 2)) / j))) And &HFFFF&
    End If
End Function

' Returns sqrt(n), rounded down. n: uint32_t.
Public Function iSqrt(ByVal n As Long) As Long
    Dim nd As Double, r As Double
    nd = U32(n)
    r = Int(Sqr(nd))         ' Calculate square root, rounded down. Excess precision does not change the result.

    ' Check that we got the right result.
    If Not (r * r - nd <= 0 And (r + 1) * (r + 1) - nd > 0) Then
        ASSERT_FAILURE "(int32_t)(r * r - n) <= 0 && (int32_t)((r + 1) * (r + 1) - n) > 0", "lib/framework/trig.cpp:160", "iSqrt", "Too badly broken sqrt function, iSqrt(%u) = %u.", nd, r
    End If

    iSqrt = FromU32(r)
End Function

' (VB) i64Sqrt of a Double below 2^53, without Variants (iHypot's case, on the path finder's hot path)
Private Function i64SqrtD(ByVal n As Double) As Long
    Dim r As Double
    r = Int(Sqr(n))
    ' Correct for different rounding & precision
    If r > 4294967295# Then r = 4294967295#
    Do While r * r > n
        r = r - 1
    Loop
    Do While n - r * r > r * 2
        r = r + 1
    Loop

    ' Check that we got the right result.
    If Not (r * r <= n And (r + 1) * (r + 1) > n) Then
        ASSERT_FAILURE "(int64_t)(r * r - n) <= 0 && (int64_t)((r + 1) * (r + 1) - n) > 0", "lib/framework/trig.cpp:187", "i64Sqrt", "Too badly broken sqrt function, i64Sqrt(%u) = %u.", CDec(n), CDec(r)
    End If

    If r < 2147483648# Then i64SqrtD = r Else i64SqrtD = FromU32(r)
End Function

' Returns sqrt(n), rounded down. n: uint64_t, a Double below 2^53 or a Decimal.
Public Function i64Sqrt(ByVal n As Variant) As Long
    Dim r As Variant, nn As Variant
    If VarType(n) <> vbDecimal And CDbl(n) < TWO_53 Then
        i64Sqrt = i64SqrtD(CDbl(n))
        Exit Function
    End If
    nn = CDec(n)
    r = CDec(Int(Sqr(CDbl(nn))))
    ' Correct for different rounding & precision
    If r > 4294967295# Then r = 4294967295#
    r = CDec(r)
    Do While r * r > nn
        r = r - 1
    Loop
    Do While nn - r * r > r * 2
        r = r + 1
    Loop

    ' Check that we got the right result.
    If Not (r * r <= nn And (r + 1) * (r + 1) > nn) Then
        ASSERT_FAILURE "(int64_t)(r * r - n) <= 0 && (int64_t)((r + 1) * (r + 1) - n) > 0", "lib/framework/trig.cpp:187", "i64Sqrt", "Too badly broken sqrt function, i64Sqrt(%u) = %u.", CDec(nn), CDec(r)
    End If

    i64Sqrt = FromU32(CDbl(r))
End Function

' Returns sqrt(x^2 + y^2), rounded down.
Public Function iHypot(ByVal x As Long, ByVal y As Long) As Long
    Dim n As Double
    n = CDbl(x) * x + CDbl(y) * y
    If n < TWO_53 Then
        iHypot = i64SqrtD(n)
    Else
        iHypot = i64Sqrt(CDec(x) * x + CDec(y) * y)
    End If
End Function

' Returns sqrt(x^2 + y^2 + z^2), rounded down. In case of overflow, returns correct result cast to (int32_t).
Public Function iHypot3(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    Dim n As Variant
    n = CDbl(x) * x + CDbl(y) * y + CDbl(z) * z
    If n >= TWO_53 Then
        n = CDec(x) * x + CDec(y) * y + CDec(z) * z
        If n >= CDec(TWO_32) * CDec(TWO_32) Then n = n - CDec(TWO_32) * CDec(TWO_32)   ' uint64_t wraps
    End If
    iHypot3 = i64Sqrt(n)
End Function

' trig.h: (int16_t)a, the wrapping intended
Public Function angleDelta(ByVal a As Long) As Long
    angleDelta = ToInt16(a)
End Function
