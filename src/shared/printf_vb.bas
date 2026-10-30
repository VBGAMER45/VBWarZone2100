Attribute VB_Name = "printf_vb"
' C's printf family for the port (no C++ counterpart). vsprintf, va and their
' helpers are QuakeVB's (src/common.bas), unchanged: checked against C's
' printf there (tests/printf_cases.bas, 300 cases).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' C: va (format, ...). A VB format argument list is passed on as a Variant
' holding the ParamArray: Dim a As Variant: a = args.
Public Function va(ByVal fmt As String, ParamArray args() As Variant) As String
    Dim a As Variant
    a = args
    va = vsprintf(fmt, a)
End Function

' (WarzoneVB) the digits of a whole Double below 2^53 (CStr and CDec keep 15 significant digits)
Private Function DoubleIntStr(ByVal x As Double) As String
    Dim hi As Double, lo As Double
    If x < 1E+15 Then
        DoubleIntStr = CStr(CDec(x))
    Else
        hi = Int(x / 100000000#)
        lo = x - hi * 100000000#
        DoubleIntStr = CStr(CDec(hi)) & Right$("0000000" & CStr(CDec(lo)), 8)
    End If
End Function

Private Function PadField(ByVal s As String, ByVal width As Long, ByVal leftjust As Boolean, ByVal zeropad As Boolean) As String
    Dim signch As String
    If Len(s) >= width Then
        PadField = s
    ElseIf leftjust Then
        PadField = s & Space$(width - Len(s))
    ElseIf zeropad Then
        If Left$(s, 1) = "-" Or Left$(s, 1) = "+" Or Left$(s, 1) = " " Then
            signch = Left$(s, 1)
            s = Mid$(s, 2)
        End If
        PadField = signch & String$(width - Len(s) - Len(signch), "0") & s
    Else
        PadField = Space$(width - Len(s)) & s
    End If
End Function

' |x| with prec decimals, rounded half to even like the C library
Private Function FmtFixed(ByVal x As Double, ByVal prec As Long) As String
    Dim scaled As Double, r As Double, f As Double, s As String
    If prec > 15 Then prec = 15
    scaled = x * (10# ^ prec)
    If scaled >= 1E+27 Then
        FmtFixed = Trim$(Str$(x))
        Exit Function
    End If
    r = Int(scaled)
    f = scaled - r
    If f > 0.5 Then
        r = r + 1
    ElseIf f = 0.5 Then
        If r - 2 * Int(r / 2) <> 0 Then r = r + 1
    End If
    s = CStr(CDec(r))
    If prec > 0 Then
        If Len(s) < prec + 1 Then s = String$(prec + 1 - Len(s), "0") & s
        s = Left$(s, Len(s) - prec) & "." & Right$(s, prec)
    End If
    FmtFixed = s
End Function

' |x| as d.ddde+XX
Private Function FmtExp(ByVal x As Double, ByVal prec As Long, ByVal upper As Boolean) As String
    Dim e As Long, m As String, es As String
    If x <> 0 Then
        e = Int(Log(x) / Log(10#))
        If x / (10# ^ e) >= 10 Then e = e + 1
        If x / (10# ^ e) < 1 Then e = e - 1
    End If
    m = FmtFixed(x / (10# ^ e), prec)
    If Left$(m, 2) = "10" Then              ' rounding carried into a new digit
        e = e + 1
        m = FmtFixed(x / (10# ^ e), prec)
    End If
    If e < 0 Then es = "-" Else es = "+"
    es = es & Format$(Abs(e), "00")
    If upper Then FmtExp = m & "E" & es Else FmtExp = m & "e" & es
End Function

Private Function StripZeros(ByVal s As String) As String
    Dim epos As Long, mant As String, tail As String
    epos = InStr(s, "e")
    If epos = 0 Then epos = InStr(s, "E")
    If epos > 0 Then
        mant = Left$(s, epos - 1)
        tail = Mid$(s, epos)
    Else
        mant = s
    End If
    If InStr(mant, ".") > 0 Then
        Do While Right$(mant, 1) = "0"
            mant = Left$(mant, Len(mant) - 1)
        Loop
        If Right$(mant, 1) = "." Then mant = Left$(mant, Len(mant) - 1)
    End If
    StripZeros = mant & tail
End Function

' Interprets a C format string: escapes, then %[flags][width][.prec][hlLzjt]conv
' with d i u x X o c s f e E g G p %.
Public Function vsprintf(ByVal fmt As String, args As Variant) As String
    Dim out As String, i As Long, n As Long, c As Long, k As Long
    Dim argi As Long, nargs As Long
    Dim leftjust As Boolean, zeropad As Boolean, plus As Boolean, space_ As Boolean, alt As Boolean
    Dim width As Long, prec As Long, s As String, d As Double, l As Long, neg As Boolean, wide As Boolean
    Dim a As Variant

    If IsArray(args) Then nargs = UBound(args) + 1 Else nargs = 0
    n = Len(fmt)
    i = 1
    Do While i <= n
        c = AscW(Mid$(fmt, i, 1))
        If c = 92 And i < n Then                          ' backslash escape
            i = i + 1
            c = AscW(Mid$(fmt, i, 1))
            Select Case c
            Case 110: out = out & vbLf                      ' \n
            Case 116: out = out & vbTab                     ' \t
            Case 114: out = out & vbCr                      ' \r
            Case 48 To 55                                   ' \ooo
                k = 0
                l = 0
                Do While k < 3 And i <= n
                    c = AscW(Mid$(fmt, i, 1))
                    If c < 48 Or c > 55 Then Exit Do
                    l = l * 8 + c - 48
                    i = i + 1
                    k = k + 1
                Loop
                i = i - 1
                out = out & ChrW$(l And 255)
            Case Else: out = out & ChrW$(c)                 ' \\ \' ...
            End Select
            i = i + 1
        ElseIf c = 37 And i < n Then                      ' %
            i = i + 1
            leftjust = False: zeropad = False: plus = False: space_ = False: alt = False
            Do While i <= n
                c = AscW(Mid$(fmt, i, 1))
                If c = 45 Then
                    leftjust = True
                ElseIf c = 48 Then
                    zeropad = True
                ElseIf c = 43 Then
                    plus = True
                ElseIf c = 32 Then
                    space_ = True
                ElseIf c = 35 Then
                    alt = True
                Else
                    Exit Do
                End If
                i = i + 1
            Loop
            width = 0
            If c = 42 Then                                  ' *
                If argi < nargs Then width = CLng(args(argi))
                argi = argi + 1
                i = i + 1
            Else
                Do While i <= n
                    c = AscW(Mid$(fmt, i, 1))
                    If c < 48 Or c > 57 Then Exit Do
                    width = width * 10 + c - 48
                    i = i + 1
                Loop
            End If
            prec = -1
            If i <= n Then
                If AscW(Mid$(fmt, i, 1)) = 46 Then
                    i = i + 1
                    prec = 0
                    If i <= n Then
                        If AscW(Mid$(fmt, i, 1)) = 42 Then
                            If argi < nargs Then prec = CLng(args(argi))
                            argi = argi + 1
                            i = i + 1
                        End If
                    End If
                    Do While i <= n
                        c = AscW(Mid$(fmt, i, 1))
                        If c < 48 Or c > 57 Then Exit Do
                        prec = prec * 10 + c - 48
                        i = i + 1
                    Loop
                End If
            End If
            wide = False
            Do While i <= n                                 ' length modifiers
                c = AscW(Mid$(fmt, i, 1))
                If c <> 104 And c <> 108 And c <> 76 And c <> 122 And c <> 106 And c <> 116 Then Exit Do   ' h l L z j t
                If c = 108 Or c = 106 Then
                    If c = 106 Or AscW(Mid$(fmt, i + 1, 1) & " ") = 108 Then wide = True   ' ll, j: a 64-bit value
                End If
                i = i + 1
            Loop
            If i > n Then Exit Do
            c = AscW(Mid$(fmt, i, 1))
            i = i + 1

            If c = 37 Then
                out = out & "%"
            Else
                If argi < nargs Then a = args(argi) Else a = Empty
                argi = argi + 1
                Select Case c
                Case 100, 105, 117                          ' d i u
                    If VarType(a) = vbDecimal Then
                        ' (WarzoneVB) a 64-bit value (PRIu64, %lld): printed in full
                        neg = (a < 0)
                        s = CStr(Abs(Fix(a)))
                    ElseIf wide And VarType(a) = vbDouble Then
                        ' (WarzoneVB) an int64_t kept in a Double (exact below 2^53): printed in full
                        d = Fix(a)
                        neg = (d < 0)
                        s = DoubleIntStr(Abs(d))
                    Else
                        Select Case VarType(a)
                        Case vbSingle, vbDouble, vbCurrency: d = Fix(a)
                        Case vbBoolean: d = -CLng(a)
                        Case vbString: d = Val(a)
                        Case vbEmpty: d = 0
                        Case Else: d = CDbl(a)
                        End Select
                        If d > 2147483647# Or d < -2147483648# Then d = 0
                        l = d
                        If c = 117 And l < 0 Then d = CDbl(l) + 4294967296# Else d = l
                        neg = (d < 0)
                        s = CStr(CDec(Abs(d)))
                    End If
                    If prec >= 0 Then
                        If prec = 0 And s = "0" Then s = ""     ' (WarzoneVB) C: a zero precision prints no digits for 0 ("%c%.0d")
                        If Len(s) < prec Then s = String$(prec - Len(s), "0") & s
                        zeropad = False
                    End If
                    If neg Then
                        s = "-" & s
                    ElseIf plus Then
                        s = "+" & s
                    ElseIf space_ Then
                        s = " " & s
                    End If
                    out = out & PadField(s, width, leftjust, zeropad)
                Case 120, 88, 111, 112                      ' x X o p
                    Select Case VarType(a)
                    Case vbSingle, vbDouble, vbCurrency, vbDecimal: l = Fix(a)
                    Case vbBoolean: l = -CLng(a)
                    Case vbString: l = Val(a)
                    Case vbEmpty: l = 0
                    Case Else: l = CLng(a)
                    End Select
                    If c = 111 Then s = LCase$(Oct$(l)) Else s = LCase$(Hex$(l))
                    If c = 88 Then s = UCase$(s)
                    If c = 112 Then s = String$(8 - Len(s), "0") & UCase$(s)
                    If prec >= 0 Then
                        If Len(s) < prec Then s = String$(prec - Len(s), "0") & s
                    End If
                    If alt And l <> 0 Then
                        If c = 120 Then s = "0x" & s
                        If c = 88 Then s = "0X" & s
                        If c = 111 Then s = "0" & s
                    End If
                    out = out & PadField(s, width, leftjust, zeropad)
                Case 99                                     ' c
                    If VarType(a) = vbString Then s = Left$(a, 1) Else s = ChrW$(CLng(a) And 255)
                    out = out & PadField(s, width, leftjust, False)
                Case 115                                    ' s
                    s = CStr(a & "")
                    If prec >= 0 Then s = Left$(s, prec)
                    out = out & PadField(s, width, leftjust, False)
                Case 102, 70, 101, 69, 103, 71              ' f F e E g G
                    d = CDbl(a)
                    CopyMemory l, ByVal VarPtr(d) + 4, 4   ' the sign bit, so -0 prints as -0 too
                    neg = (l < 0)
                    d = Abs(d)
                    If prec < 0 Then prec = 6
                    If c = 102 Or c = 70 Then
                        s = FmtFixed(d, prec)
                        If alt And prec = 0 Then s = s & "."
                    ElseIf c = 101 Or c = 69 Then
                        s = FmtExp(d, prec, c = 69)
                    Else
                        If prec = 0 Then prec = 1
                        ' the exponent X of the %e form after rounding decides
                        s = FmtExp(d, prec - 1, c = 71)
                        l = Val(Mid$(s, InStr(LCase$(s), "e") + 1))
                        If l >= -4 And l < prec Then s = FmtFixed(d, prec - 1 - l)
                        If Not alt Then s = StripZeros(s)
                    End If
                    If neg Then
                        s = "-" & s
                    ElseIf plus Then
                        s = "+" & s
                    ElseIf space_ Then
                        s = " " & s
                    End If
                    out = out & PadField(s, width, leftjust, zeropad)
                Case Else
                    out = out & "%" & ChrW$(c)
                End Select
            End If
        Else
            out = out & ChrW$(c)
            i = i + 1
        End If
    Loop
    vsprintf = out
End Function
