Attribute VB_Name = "cstdlib_vb"
' C's <stdlib.h> for the port (no C++ counterpart; compiled into every
' project): strtol, strtoul, strtoll, strtod and atoi with C's rules (C
' locale), and exit. endIdx gets the count of characters used (C: end -
' str): 0 when nothing could be converted. Also C++'s std::stoi, stoul,
' stoull, stod and stof (<string>): False and the exception's what() (the
' Microsoft library's texts) where they throw. rand and srand are the
' Microsoft runtime's generator; each project has its own state (C: one
' per process).
Option Explicit

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Private holdrand As Long
Private holdrandSet As Boolean          ' holdrand starts at 1

Public Sub srand(ByVal seed As Long)
    holdrand = seed
    holdrandSet = True
End Sub

Public Function rand() As Long
    If Not holdrandSet Then
        holdrand = 1
        holdrandSet = True
    End If
    holdrand = holdrand * 214013 + 2531011
    rand = SHRU(holdrand, 16) And &H7FFF&
End Function

Private Function IsCSpace(ByVal c As Long) As Boolean
    IsCSpace = (c = 32 Or (c >= 9 And c <= 13))
End Function

Private Function DigitValue(ByVal c As Long) As Long
    If c >= 48 And c <= 57 Then
        DigitValue = c - 48
    ElseIf c >= 97 And c <= 122 Then
        DigitValue = c - 87
    ElseIf c >= 65 And c <= 90 Then
        DigitValue = c - 55
    Else
        DigitValue = 99
    End If
End Function

' The common part of the integer parsers: the magnitude (a Double, saturated at limit + 1),
' the sign, and the characters used.
Private Function ParseInteger(ByVal s As String, ByVal base As Long, neg As Boolean, endIdx As Long, ByVal limit As Double) As Double
    Dim i As Long, n As Long, c As Long, d As Long, v As Double, any_ As Boolean
    n = Len(s)
    i = 1
    Do While i <= n
        If Not IsCSpace(AscW(Mid$(s, i, 1))) Then Exit Do
        i = i + 1
    Loop
    neg = False
    If i <= n Then
        c = AscW(Mid$(s, i, 1))
        If c = 45 Or c = 43 Then
            neg = (c = 45)
            i = i + 1
        End If
    End If
    If (base = 0 Or base = 16) And i + 1 <= n Then
        If Mid$(s, i, 1) = "0" And LCase$(Mid$(s, i + 1, 1)) = "x" Then
            If i + 2 <= n Then
                If DigitValue(AscW(Mid$(s, i + 2, 1))) < 16 Then
                    base = 16
                    i = i + 2
                End If
            End If
        End If
    End If
    If base = 0 Then
        If i <= n Then
            If Mid$(s, i, 1) = "0" Then base = 8 Else base = 10
        Else
            base = 10
        End If
    End If
    Do While i <= n
        d = DigitValue(AscW(Mid$(s, i, 1)))
        If d >= base Then Exit Do
        any_ = True
        If v <= limit Then v = v * base + d
        i = i + 1
    Loop
    If any_ Then endIdx = i - 1 Else endIdx = 0
    If v > limit + 1 Then v = limit + 1
    ParseInteger = v
End Function

' C: strtol (long is 32 bits on Windows): clamps to LONG_MIN / LONG_MAX on overflow.
Public Function strtol(ByVal s As String, endIdx As Long, ByVal base As Long) As Long
    Dim neg As Boolean, v As Double
    v = ParseInteger(s, base, neg, endIdx, 2147483647#)
    If neg Then
        If v > 2147483648# Then v = 2147483648#
        strtol = FromU32(-v)
    Else
        If v > 2147483647# Then v = 2147483647#
        strtol = v
    End If
End Function

' C: strtoul: the uint32_t bit pattern; ULONG_MAX on overflow; "-n" is the unsigned negation.
Public Function strtoul(ByVal s As String, endIdx As Long, ByVal base As Long) As Long
    Dim neg As Boolean, v As Double
    v = ParseInteger(s, base, neg, endIdx, 4294967295#)
    If v > 4294967295# Then
        strtoul = -1
    ElseIf neg Then
        strtoul = FromU32(-v)
    Else
        strtoul = FromU32(v)
    End If
End Function

' C: strtoll, as a Double (exact below 2^53); clamps at the int64_t limits.
Public Function strtoll(ByVal s As String, endIdx As Long, ByVal base As Long) As Double
    Dim neg As Boolean, v As Double
    v = ParseInteger(s, base, neg, endIdx, 9.22337203685478E+18)
    If neg Then strtoll = -v Else strtoll = v
End Function

' C: strtod: [space][sign](digits[.digits] | .digits)[(e|E)[sign]digits], inf, infinity, nan.
Public Function strtod(ByVal s As String, endIdx As Long) As Double
    Dim i As Long, n As Long, st As Long, c As String, digits As Long, mant As String
    Dim expStart As Long, neg As Boolean, rest As String
    n = Len(s)
    i = 1
    Do While i <= n
        If Not IsCSpace(AscW(Mid$(s, i, 1))) Then Exit Do
        i = i + 1
    Loop
    st = i
    If i <= n Then
        c = Mid$(s, i, 1)
        If c = "-" Or c = "+" Then
            neg = (c = "-")
            i = i + 1
        End If
    End If
    rest = LCase$(Mid$(s, i, 8))
    If Left$(rest, 3) = "inf" Then
        If Left$(rest, 8) = "infinity" Then endIdx = i + 7 Else endIdx = i + 2
        ' VB cannot write an infinity literal: the largest Double stands in
        If neg Then strtod = -1.79769313486231E+308 Else strtod = 1.79769313486231E+308
        Exit Function
    End If
    Do While i <= n
        c = Mid$(s, i, 1)
        If c < "0" Or c > "9" Then Exit Do
        digits = digits + 1
        i = i + 1
    Loop
    If i <= n Then
        If Mid$(s, i, 1) = "." Then
            i = i + 1
            Do While i <= n
                c = Mid$(s, i, 1)
                If c < "0" Or c > "9" Then Exit Do
                digits = digits + 1
                i = i + 1
            Loop
        End If
    End If
    If digits = 0 Then
        endIdx = 0
        Exit Function
    End If
    mant = Mid$(s, st, i - st)
    If i <= n Then
        If LCase$(Mid$(s, i, 1)) = "e" Then
            expStart = i
            i = i + 1
            If i <= n Then
                c = Mid$(s, i, 1)
                If c = "-" Or c = "+" Then i = i + 1
            End If
            digits = 0
            Do While i <= n
                c = Mid$(s, i, 1)
                If c < "0" Or c > "9" Then Exit Do
                digits = digits + 1
                i = i + 1
            Loop
            If digits = 0 Then
                i = expStart                    ' "1e" is 1 followed by "e"
            End If
        End If
    End If
    endIdx = i - 1
    strtod = Val(Mid$(s, st, i - st))           ' Val reads C's notation in any locale
End Function

' C++: std::stoi (base 10): invalid_argument, or out_of_range beyond int
Public Function std_stoi(ByVal s As String, value As Long, errText As String) As Boolean
    Dim endIdx As Long, d As Double
    d = strtoll(s, endIdx, 10)
    If endIdx = 0 Then
        errText = "invalid stoi argument"
    ElseIf d > 2147483647# Or d < -2147483648# Then
        errText = "stoi argument out of range"
    Else
        value = d
        std_stoi = True
    End If
End Function

' C++: std::stoul (unsigned long is 32 bits on Windows; "-n" wraps as strtoul does): the bits
Public Function std_stoul(ByVal s As String, value As Long, errText As String) As Boolean
    Dim endIdx As Long, d As Double
    d = strtoll(s, endIdx, 10)
    If endIdx = 0 Then
        errText = "invalid stoul argument"
    ElseIf Abs(d) > 4294967295# Then
        errText = "stoul argument out of range"
    Else
        If d < 0 Then d = d + 4294967296#
        value = FromU32(d)
        std_stoul = True
    End If
End Function

' C++: std::stoull, as a Double (exact to 2^53)
Public Function std_stoull(ByVal s As String, value As Double, errText As String) As Boolean
    Dim endIdx As Long, d As Double
    d = strtoll(s, endIdx, 10)
    If endIdx = 0 Then
        errText = "invalid stoull argument"
    ElseIf Abs(d) > 1.84467440737096E+19 Then
        errText = "stoull argument out of range"
    Else
        If d < 0 Then d = d + 1.84467440737096E+19
        value = d
        std_stoull = True
    End If
End Function

' C++: std::stod: invalid_argument, or out_of_range on overflow
Public Function std_stod(ByVal s As String, value As Double, errText As String) As Boolean
    Dim endIdx As Long, d As Double
    On Error GoTo overflow
    d = strtod(s, endIdx)
    On Error GoTo 0
    If endIdx = 0 Then
        errText = "invalid stod argument"
    Else
        value = d
        std_stod = True
    End If
    Exit Function
overflow:
    errText = "stod argument out of range"
End Function

' C++: std::stof: invalid_argument, or out_of_range beyond FLT_MAX
Public Function std_stof(ByVal s As String, value As Single, errText As String) As Boolean
    Dim endIdx As Long, d As Double
    On Error GoTo overflow
    d = strtod(s, endIdx)
    On Error GoTo 0
    If endIdx = 0 Then
        errText = "invalid stof argument"
    ElseIf Abs(d) > 3.40282346638529E+38 Then
        errText = "stof argument out of range"
    Else
        value = d
        std_stof = True
    End If
    Exit Function
overflow:
    errText = "stof argument out of range"
End Function

' C: atoi (Microsoft's: strtol's result, so out of range clamps)
Public Function atoi(ByVal s As String) As Long
    Dim endIdx As Long
    atoi = strtol(s, endIdx, 10)
End Function

' C: exit (status). The port's C streams are unbuffered, so there is
' nothing to flush; the process ends at once, without unwinding, as in C.
Public Sub exit_(ByVal status As Long)
    ExitProcess status
End Sub
