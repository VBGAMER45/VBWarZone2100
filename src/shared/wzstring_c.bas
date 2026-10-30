Attribute VB_Name = "wzstring_c"
' lib/framework/wzstring.cpp: the WzString methods whose behaviour VB's
' own string functions do not give (a WzString is a VB String). toLower
' changes only A-Z, as std::tolower in the classic locale does on the
' UTF-8 bytes; trimmed removes the classic locale's spaces (" \f\n\r\t\v");
' split follows WzString::split code point by code point (a VB String's
' UTF-16 units here). Stateless, so compiled into every project.
Option Explicit

' WzString::toLower
Public Function WzString_toLower(ByVal s As String) As String
    Dim i As Long, c As Long
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c >= 65 And c <= 90 Then Mid$(s, i, 1) = ChrW$(c + 32)
    Next
    WzString_toLower = s
End Function

Private Function IsClassicSpace(ByVal c As Long) As Boolean
    IsClassicSpace = (c = 32 Or (c >= 9 And c <= 13))
End Function

' WzString::trimmed (classic locale)
Public Function WzString_trimmed(ByVal s As String) As String
    Dim a As Long, b As Long
    a = 1
    b = Len(s)
    Do While a <= b
        If Not IsClassicSpace(AscW(Mid$(s, a, 1))) Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If Not IsClassicSpace(AscW(Mid$(s, b, 1))) Then Exit Do
        b = b - 1
    Loop
    WzString_trimmed = Mid$(s, a, b - a + 1)
End Function

' WzString::split: the tokens in tokens(0 .. count - 1); returns the count
Public Function WzString_split(ByVal s As String, ByVal delimiter As String, tokens() As String) As Long
    Dim n As Long, i As Long, tokenStart As Long, cur As Long, dLen As Long
    Dim lastCodepointWasDelimiterEnd As Boolean
    dLen = Len(delimiter)
    Erase tokens
    n = 0
    cur = 1                             ' currentDelimiterCodepoint (1-based in delimiter)
    tokenStart = 1                      ' it_currenttokenstart
    For i = 1 To Len(s)
        If dLen > 0 Then
            If Mid$(s, i, 1) <> Mid$(delimiter, cur, 1) Then
                cur = 1
                lastCodepointWasDelimiterEnd = False
            Else
                cur = cur + 1
            End If
        End If

        If cur = dLen + 1 Then
            ReDim Preserve tokens(0 To n)
            tokens(n) = Mid$(s, tokenStart, i + 1 - dLen - tokenStart)
            n = n + 1
            tokenStart = i + 1
            cur = 1
            lastCodepointWasDelimiterEnd = True
        End If
    Next

    If tokenStart <= Len(s) Then
        ReDim Preserve tokens(0 To n)
        tokens(n) = Mid$(s, tokenStart)
        n = n + 1
    ElseIf lastCodepointWasDelimiterEnd And dLen > 0 Then
        ' string ends with delimiter (and delimiter is not an empty string)
        ' append an empty string to the end of the list
        ReDim Preserve tokens(0 To n)
        tokens(n) = ""
        n = n + 1
    End If
    WzString_split = n
End Function
