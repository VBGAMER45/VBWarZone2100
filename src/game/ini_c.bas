Attribute VB_Name = "ini_c"
' 3rdparty/inih/ini.c, built as Warzone builds it (3rdparty/CMakeLists.txt):
' INI_ALLOW_MULTILINE=0, INI_ALLOW_BOM=1, INI_ALLOW_INLINE_COMMENTS=0,
' INI_MAX_LINE=1024 on the stack (so INI_ALLOW_REALLOC has no effect: a
' longer line is read in 1023-byte pieces, each parsed as a line), and the
' defaults for the rest. The text is UTF-8 bytes, as in C; section, name and
' value reach the handler as Strings. ini_parse_string's handler is always
' INIReaderWriter::ValueHandler, so it is called directly. Also here:
' INIDetail::ascii_strcasecmp (INIReaderWriter.cpp).
Option Explicit

Private Const INI_MAX_LINE As Long = 1024
Private Const MAX_SECTION As Long = 50
Private Const MAX_NAME As Long = 50

Private Function fixed_isspace(ByVal ch As Long) As Boolean
    Select Case ch
    Case 32, 9, 10, 13, 12, 11
        fixed_isspace = True
    End Select
End Function

' The String of line(first .. first + count - 1), at most limit - 1 bytes (strncpy0 into char[limit]).
Private Function LineText(lineBuf() As Byte, ByVal first As Long, ByVal count As Long, ByVal limit As Long) As String
    If limit > 0 And count > limit - 1 Then count = limit - 1
    LineText = utf8_decode(lineBuf, first, count)
End Function

' C: ini_parse_string (string, handler, user): s(0 .. length - 1) is the text (length = strlen).
' Returns 0 on success, else the line number of the first error.
Public Function ini_parse_string(s() As Byte, ByVal length As Long, ByVal user As INIReaderWriter) As Long
    Dim lineBuf(0 To INI_MAX_LINE - 1) As Byte, lineEnd As Long
    Dim ctx_ptr As Long, ctx_num_left As Long, num As Long, c As Byte
    Dim section As String, start As Long, e As Long, nameEnd As Long, valueStart As Long
    Dim lineno As Long, error_ As Long, ch As Long

    ctx_ptr = 0
    ctx_num_left = length
    ' Scan through stream line by line
    Do
        ' ini_reader_string (line, INI_MAX_LINE, stream)
        If ctx_num_left = 0 Then Exit Do
        num = INI_MAX_LINE
        lineEnd = 0
        Do While num > 1 And ctx_num_left <> 0
            c = s(ctx_ptr)
            ctx_ptr = ctx_ptr + 1
            ctx_num_left = ctx_num_left - 1
            lineBuf(lineEnd) = c
            lineEnd = lineEnd + 1
            If c = 10 Then Exit Do
            num = num - 1
        Loop

        lineno = lineno + 1

        start = 0
        If lineno = 1 And lineEnd >= 3 Then
            If lineBuf(0) = &HEF And lineBuf(1) = &HBB And lineBuf(2) = &HBF Then start = 3
        End If
        ' start = lskip(rstrip(start))
        Do While lineEnd > start
            If Not fixed_isspace(lineBuf(lineEnd - 1)) Then Exit Do
            lineEnd = lineEnd - 1
        Loop
        Do While start < lineEnd
            If Not fixed_isspace(lineBuf(start)) Then Exit Do
            start = start + 1
        Loop

        If start = lineEnd Then
            ch = 0
        Else
            ch = lineBuf(start)
        End If
        If ch = 0 Or ch = 59 Or ch = 35 Then
            ' Start-of-line comment (strchr(";#", *start) also matches the terminator: an empty line)
        ElseIf ch = 91 Then
            ' A "[section]" line
            e = start + 1
            Do While e < lineEnd
                If lineBuf(e) = 93 Then Exit Do
                e = e + 1
            Loop
            If e < lineEnd Then
                section = LineText(lineBuf, start + 1, e - (start + 1), MAX_SECTION)
            ElseIf error_ = 0 Then
                ' No ']' found on section line
                error_ = lineno
            End If
        Else
            ' Not a comment, must be a name[=:]value pair
            e = start
            Do While e < lineEnd
                If lineBuf(e) = 61 Or lineBuf(e) = 58 Then Exit Do
                e = e + 1
            Loop
            If e < lineEnd Then
                nameEnd = e
                Do While nameEnd > start
                    If Not fixed_isspace(lineBuf(nameEnd - 1)) Then Exit Do
                    nameEnd = nameEnd - 1
                Loop
                valueStart = e + 1
                Do While valueStart < lineEnd
                    If Not fixed_isspace(lineBuf(valueStart)) Then Exit Do
                    valueStart = valueStart + 1
                Loop
                ' Valid name[=:]value pair found, call handler
                If user.ValueHandler(section, LineText(lineBuf, start, nameEnd - start, 0), LineText(lineBuf, valueStart, lineEnd - valueStart, 0)) = 0 And error_ = 0 Then
                    error_ = lineno
                End If
            ElseIf error_ = 0 Then
                ' No '=' or ':' found on name[=:]value line
                error_ = lineno
            End If
        End If
    Loop

    ini_parse_string = error_
End Function

' INIDetail::ascii_strcasecmp: <0, 0, >0 like strcasecmp, folding only A-Z.
Public Function ascii_strcasecmp(ByVal s1 As String, ByVal s2 As String) As Long
    Dim i As Long, n As Long, c1 As Long, c2 As Long
    n = Len(s1)
    If Len(s2) < n Then n = Len(s2)
    For i = 1 To n
        c1 = AscW(Mid$(s1, i, 1)) And &HFFFF&
        c2 = AscW(Mid$(s2, i, 1)) And &HFFFF&
        If c1 >= 65 And c1 <= 90 Then c1 = c1 + 32
        If c2 >= 65 And c2 <= 90 Then c2 = c2 + 32
        If c1 <> c2 Then
            ascii_strcasecmp = c1 - c2
            Exit Function
        End If
    Next
    ascii_strcasecmp = Len(s1) - Len(s2)
End Function
