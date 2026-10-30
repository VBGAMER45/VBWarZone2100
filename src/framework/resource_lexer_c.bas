Attribute VB_Name = "resource_lexer_c"
' lib/framework/resource_lexer.lpp: the flex lexer for resource (*.wrf)
' files (prefix res_), written out by hand with flex's rules: the longest
' match wins, the earlier rule on a tie; start conditions COMMENT, QUOTE,
' SLCOMMENT; yylineno counts the newlines matched; yytext is the last
' match; res_lex_destroy resets the scanner (yylineno 1, INITIAL).
' YY_INPUT is lib/framework/lexer_input.cpp's LEXINPUT_PHYSFS: the file
' (yyextra, the PhysFS handle) is read whole on the first call.
' The tokens are resource_parser.ypp's, as bison numbers them (from 258);
' any other character is returned as its char (signed) value.
' res_lval.sval is res_lval_sval.
Option Explicit

' resource_parser.h: the token numbers
Public Const TEXT_T As Long = 258
Public Const QTEXT_T As Long = 259          ' Text with double quotes surrounding it
Public Const DIRECTORY As Long = 260
Public Const FILETOKEN As Long = 261

' start conditions
Private Const SC_INITIAL As Long = 0
Private Const SC_COMMENT As Long = 1
Private Const SC_QUOTE As Long = 2
Private Const SC_SLCOMMENT As Long = 3

Public res_lval_sval As String          ' res_lval.sval

Private pBuffer As Long                 ' YY_EXTRA_TYPE: the input (a PhysFS file handle)
Private yy_init As Boolean              ' the input has been read
Private inBuf() As Byte
Private inLen As Long
Private inPos As Long
Private yylineno As Long
Private yytext As String
Private yy_start As Long

Public Sub res_set_extra(ByVal user_defined As Long)
    pBuffer = user_defined
End Sub

Public Function res_get_extra() As Long
    res_get_extra = pBuffer
End Function

Public Function res_get_lineno() As Long
    If Not yy_init Then res_get_lineno = 1 Else res_get_lineno = yylineno
End Function

Public Function res_get_text() As String
    res_get_text = yytext
End Function

' For non-reentrant C scanner only.
Public Function res_lex_destroy() As Long
    Erase inBuf
    inLen = 0
    inPos = 0
    yy_init = False
    yytext = ""
    res_lex_destroy = 0
End Function

' the scanner's first call: yylineno 1, INITIAL, and the input (lexer_input, LEXINPUT_PHYSFS)
Private Sub InitScanner()
    Dim length As Double, got As Long
    yylineno = 1
    yy_start = SC_INITIAL
    inLen = 0
    inPos = 0
    length = PHYSFS_fileLength(pBuffer)
    If length > 0 Then
        ReDim inBuf(0 To CLng(length) - 1)
        got = PHYSFS_readBytes(pBuffer, inBuf, 0, CLng(length))
        If got > 0 Then inLen = got
    End If
    yy_init = True
End Sub

Private Function IsAlpha(ByVal c As Long) As Boolean
    IsAlpha = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

' [-0-9_a-zA-Z]
Private Function IsIdentChar(ByVal c As Long) As Boolean
    IsIdentChar = IsAlpha(c) Or (c >= 48 And c <= 57) Or c = 45 Or c = 95
End Function

' the byte at inPos + k, or -1 past the end
Private Function PeekByte(ByVal k As Long) As Long
    If inPos + k < inLen Then PeekByte = inBuf(inPos + k) Else PeekByte = -1
End Function

' match n bytes: yytext, and yylineno for their newlines
Private Sub Match(ByVal n As Long)
    Dim i As Long
    yytext = utf8_decode(inBuf, inPos, n)
    For i = inPos To inPos + n - 1
        If inBuf(i) = 10 Then yylineno = yylineno + 1
    Next
    inPos = inPos + n
End Sub

Public Function res_lex() As Long
    Dim c As Long, n As Long
    If Not yy_init Then InitScanner
    Do
        If inPos >= inLen Then
            ' <<EOF>>: yyterminate()
            yytext = ""
            res_lex = 0
            Exit Function
        End If
        c = inBuf(inPos)
        Select Case yy_start
        Case SC_INITIAL
            If IsAlpha(c) Then
                ' Match to key words / Match text values: [a-zA-Z][-0-9_a-zA-Z]*
                n = 1
                Do While inPos + n < inLen
                    If Not IsIdentChar(inBuf(inPos + n)) Then Exit Do
                    n = n + 1
                Loop
                Match n
                If yytext = "directory" Then
                    res_lex = DIRECTORY
                ElseIf yytext = "file" Then
                    res_lex = FILETOKEN
                Else
                    res_lval_sval = yytext
                    res_lex = TEXT_T
                End If
                Exit Function
            ElseIf c = 34 Then
                ' Match quoted text
                If PeekByte(1) = 34 Then
                    Match 2
                    res_lval_sval = ""
                    res_lex = QTEXT_T
                    Exit Function
                End If
                Match 1
                yy_start = SC_QUOTE
            ElseIf c = 32 Or c = 9 Or c = 10 Or c = 13 Then
                ' Skip white space
                Match 1
            ElseIf c = 47 And PeekByte(1) = 42 Then
                ' Strip comments
                Match 2
                yy_start = SC_COMMENT
            ElseIf c = 47 And PeekByte(1) = 47 Then
                ' Strip single line comments
                Match 2
                yy_start = SC_SLCOMMENT
            Else
                ' Match anything that's been missed and pass it as a char
                Match 1
                If c >= 128 Then res_lex = c - 256 Else res_lex = c
                Exit Function
            End If
        Case SC_QUOTE
            If c = 34 Then
                Match 1
                yy_start = SC_INITIAL
            ElseIf c = 10 Then
                Match 1
                res_error "Unexpected end of line in string"
            Else
                n = 1
                Do While inPos + n < inLen
                    If inBuf(inPos + n) = 34 Or inBuf(inPos + n) = 10 Then Exit Do
                    n = n + 1
                Loop
                Match n
                res_lval_sval = yytext
                res_lex = QTEXT_T
                Exit Function
            End If
        Case SC_COMMENT
            If c = 42 And PeekByte(1) = 47 Then
                If PeekByte(2) = 10 Then Match 3 Else Match 2
                yy_start = SC_INITIAL
            Else
                Match 1
            End If
        Case SC_SLCOMMENT
            If c = 10 Then
                Match 1
                yy_start = SC_INITIAL
            Else
                n = 1
                Do While inPos + n < inLen
                    If inBuf(inPos + n) = 10 Then Exit Do
                    n = n + 1
                Loop
                Match n
            End If
        End Select
    Loop
End Function
