Attribute VB_Name = "test_json"
' WzTests: the JSON modes (as tools/jsonref.cpp), through WzFramework's
' nlohmann_json.
Option Explicit

Private Function ParseFile(ByVal path As String, ByVal ordered As Boolean) As nlohmann_json
    Dim b() As Byte, n As Long
    n = tests_main.ReadBytes(path, b)
    Set ParseFile = json_parse_bytes(b, 0, n, ordered)
End Function

Private Function Doubles(ByVal path As String) As String
    Dim b() As Byte, n As Long, lines() As String, i As Long, s As String, out As String
    n = tests_main.ReadBytes(path, b)
    lines = Split(Replace(StrConv(b, vbUnicode), vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        s = Trim$(lines(i))
        If Len(s) = 16 Then
            out = out & json_value(DoubleFromBits(CLng("&H" & Left$(s, 8)), CLng("&H" & Right$(s, 8)))).dump() & vbLf
        End If
    Next
    Doubles = out
End Function

Public Function RunJson(ByVal mode As String, ByVal inPath As String) As String
    Dim j As nlohmann_json
    If mode = "doubles" Then
        RunJson = Doubles(inPath)
        Exit Function
    End If
    On Error GoTo failed
    Select Case mode
    Case "dumpo"
        RunJson = ParseFile(inPath, True).dump(4) & vbLf
    Case "dump"
        RunJson = ParseFile(inPath, False).dump(4) & vbLf
    Case "compact"
        RunJson = ParseFile(inPath, False).dump()
    Case "ascii"
        RunJson = ParseFile(inPath, False).dump(-1, " ", True)
    Case "error"
        Set j = ParseFile(inPath, False)
        RunJson = "ok" & vbLf
    Case Else
        RunJson = "unknown mode " & mode & vbLf
    End Select
    Exit Function
failed:
    If mode = "error" Then
        RunJson = Err.Description & vbLf
    Else
        RunJson = "error: " & Err.Description & vbLf
    End If
End Function
