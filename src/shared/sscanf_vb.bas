Attribute VB_Name = "sscanf_vb"
' C's sscanf for the port (no C++ counterpart; compiled into every
' project): the conversions %d %i %u %x %f (e g) %s %c %n and %[set], with an
' optional * (no assignment), width, and length modifier (ignored: the
' variable passed decides the type). The results go into the variables
' given after the format (a ParamArray passes them ByRef): Long for the
' integers (%u and %x as the bits of a uint32_t), Single or Double for %f,
' String for %s, %c and %[. Returns the number of assignments, or -1 (EOF)
' when the input ends before the first conversion, as C does. %f reads
' C's syntax and converts through strtod (Double) to the variable's type.
Option Explicit

Private Function IsSpaceCh(ByVal c As Long) As Boolean
    IsSpaceCh = (c = 32 Or (c >= 9 And c <= 13))
End Function

Private Function HexVal(ByVal c As Long) As Long
    If c >= 48 And c <= 57 Then
        HexVal = c - 48
    ElseIf c >= 97 And c <= 102 Then
        HexVal = c - 87
    ElseIf c >= 65 And c <= 70 Then
        HexVal = c - 55
    Else
        HexVal = -1
    End If
End Function

' whether c is in the scanset set (the text between [ and ], without a leading ^)
Private Function InScanset(ByVal set_ As String, ByVal c As Long) As Boolean
    Dim i As Long, n As Long, a As Long, b As Long
    n = Len(set_)
    i = 1
    Do While i <= n
        a = AscW(Mid$(set_, i, 1))
        If i + 2 <= n And Mid$(set_, i + 1, 1) = "-" Then
            b = AscW(Mid$(set_, i + 2, 1))
            If c >= a And c <= b Then
                InScanset = True
                Exit Function
            End If
            i = i + 3
        Else
            If c = a Then
                InScanset = True
                Exit Function
            End If
            i = i + 1
        End If
    Loop
End Function

Public Function sscanf(ByVal s As String, ByVal fmt As String, ParamArray args() As Variant) As Long
    Dim fi As Long, fn As Long, si As Long, sn As Long, f As String, c As Long
    Dim suppress As Boolean, width As Long, conv As String, argi As Long, assigned As Long
    Dim neg As Boolean, v As Double, digits As Long, st As Long, endIdx As Long, sv As String
    Dim set_ As String, negate As Boolean, k As Long, anyConversion As Boolean

    fn = Len(fmt)
    sn = Len(s)
    fi = 1
    si = 1
    Do While fi <= fn
        f = Mid$(fmt, fi, 1)
        If IsSpaceCh(AscW(f)) Then
            ' white space in the format: any amount of white space in the input
            Do While si <= sn
                If Not IsSpaceCh(AscW(Mid$(s, si, 1))) Then Exit Do
                si = si + 1
            Loop
            fi = fi + 1
        ElseIf f <> "%" Or Mid$(fmt, fi + 1, 1) = "%" Then
            ' an ordinary character (or %%) must match
            If f = "%" Then
                fi = fi + 1
                Do While si <= sn
                    If Not IsSpaceCh(AscW(Mid$(s, si, 1))) Then Exit Do
                    si = si + 1
                Loop
            End If
            If si > sn Then GoTo input_failure
            If Mid$(s, si, 1) <> f Then GoTo done
            si = si + 1
            fi = fi + 1
        Else
            fi = fi + 1
            suppress = False
            If Mid$(fmt, fi, 1) = "*" Then
                suppress = True
                fi = fi + 1
            End If
            width = 0
            Do While fi <= fn
                c = AscW(Mid$(fmt, fi, 1))
                If c < 48 Or c > 57 Then Exit Do
                width = width * 10 + c - 48
                fi = fi + 1
            Loop
            Do While InStr("hlLqjzt", Mid$(fmt, fi, 1)) > 0 And fi <= fn
                fi = fi + 1
            Loop
            conv = Mid$(fmt, fi, 1)
            fi = fi + 1
            If conv = "[" Then
                ' the scanset, to the closing ] (a ] right after [ or [^ belongs to the set)
                negate = False
                If Mid$(fmt, fi, 1) = "^" Then
                    negate = True
                    fi = fi + 1
                End If
                k = fi
                If Mid$(fmt, fi, 1) = "]" Then fi = fi + 1
                Do While fi <= fn
                    If Mid$(fmt, fi, 1) = "]" Then Exit Do
                    fi = fi + 1
                Loop
                set_ = Mid$(fmt, k, fi - k)
                fi = fi + 1
            End If
            If conv = "n" Then
                If Not suppress Then
                    args(argi) = si - 1
                    argi = argi + 1
                End If
                GoTo next_directive
            End If
            If conv <> "c" And conv <> "[" Then
                Do While si <= sn
                    If Not IsSpaceCh(AscW(Mid$(s, si, 1))) Then Exit Do
                    si = si + 1
                Loop
            End If
            If si > sn Then GoTo input_failure
            If width = 0 Then width = &H7FFFFFFF
            Select Case conv
            Case "d", "i", "u", "x", "X"
                st = si
                neg = False
                c = AscW(Mid$(s, si, 1))
                If (c = 45 Or c = 43) And width > 0 Then
                    neg = (c = 45)
                    si = si + 1
                End If
                v = 0
                digits = 0
                If conv = "x" Or conv = "X" Then
                    If si + 1 <= sn And si - st + 2 <= width Then
                        If Mid$(s, si, 1) = "0" And LCase$(Mid$(s, si + 1, 1)) = "x" Then
                            If si + 2 <= sn Then
                                If HexVal(AscW(Mid$(s, si + 2, 1))) >= 0 Then si = si + 2
                            End If
                        End If
                    End If
                    Do While si <= sn And si - st < width
                        k = HexVal(AscW(Mid$(s, si, 1)))
                        If k < 0 Then Exit Do
                        v = v * 16 + k
                        If v >= 4294967296# Then v = v - Int(v / 4294967296#) * 4294967296#
                        digits = digits + 1
                        si = si + 1
                    Loop
                Else
                    Do While si <= sn And si - st < width
                        c = AscW(Mid$(s, si, 1))
                        If c < 48 Or c > 57 Then Exit Do
                        v = v * 10 + c - 48
                        If v >= 4294967296# Then v = v - Int(v / 4294967296#) * 4294967296#
                        digits = digits + 1
                        si = si + 1
                    Loop
                End If
                If digits = 0 Then GoTo done
                If neg Then v = -v
                v = v - Int(v / 4294967296#) * 4294967296#     ' the 32 bits
                If v >= 2147483648# Then v = v - 4294967296#
                If Not suppress Then
                    args(argi) = CLng(v)
                    argi = argi + 1
                    assigned = assigned + 1
                End If
            Case "f", "e", "g", "E", "G", "a"
                sv = Mid$(s, si, IIf(width < sn, width, sn))
                v = strtod(sv, endIdx)
                If endIdx = 0 Then GoTo done
                si = si + endIdx
                If Not suppress Then
                    args(argi) = v
                    argi = argi + 1
                    assigned = assigned + 1
                End If
            Case "s"
                st = si
                Do While si <= sn And si - st < width
                    If IsSpaceCh(AscW(Mid$(s, si, 1))) Then Exit Do
                    si = si + 1
                Loop
                If Not suppress Then
                    args(argi) = Mid$(s, st, si - st)
                    argi = argi + 1
                    assigned = assigned + 1
                End If
            Case "c"
                If width = &H7FFFFFFF Then width = 1
                If si + width - 1 > sn Then GoTo input_failure
                If Not suppress Then
                    args(argi) = Mid$(s, si, width)
                    argi = argi + 1
                    assigned = assigned + 1
                End If
                si = si + width
            Case "["
                st = si
                Do While si <= sn And si - st < width
                    If InScanset(set_, AscW(Mid$(s, si, 1))) = negate Then Exit Do
                    si = si + 1
                Loop
                If si = st Then GoTo done
                If Not suppress Then
                    args(argi) = Mid$(s, st, si - st)
                    argi = argi + 1
                    assigned = assigned + 1
                End If
            Case Else
                GoTo done
            End Select
            anyConversion = True
next_directive:
        End If
    Loop
done:
    sscanf = assigned
    Exit Function
input_failure:
    If anyConversion Then sscanf = assigned Else sscanf = -1
End Function
