Attribute VB_Name = "level_lexer_c"
' src/level_lexer.lpp: the flex lexer for level description files (prefix
' lev_), written out by hand with flex's rules: the longest match wins, the
' earlier rule on a tie (so "expand_limbo" is one keyword, "levels" an
' identifier); start conditions COMMENT, SLCOMMENT, QUOTE; yylineno counts
' the newlines matched; yytext is the last match; lev_lex_destroy resets
' the scanner. YY_INPUT is lib/framework/lexer_input.cpp's LEXINPUT_BUFFER:
' lev_set_extra takes the buffer. The tokens are src/levelint.h's; any
' other character is returned as its char (signed) value. pLevToken and
' levVal are levels_c's (levels.cpp defines them); aText is YYLMAX (255)
' chars, so a token is cut to 254.
Option Explicit

' levelint.h: return values from the lexer
Public Const LTK_LEVEL As Long = &H100         ' level key word
Public Const LTK_PLAYERS As Long = &H101       ' players key word
Public Const LTK_TYPE As Long = &H102          ' type key word
Public Const LTK_DATA As Long = &H103          ' data key word
Public Const LTK_GAME As Long = &H104          ' game key word
Public Const LTK_CAMPAIGN As Long = &H105      ' campaign key word
Public Const LTK_CAMSTART As Long = &H106      ' camstart key word
Public Const LTK_CAMCHANGE As Long = &H107     ' camchange key word
Public Const LTK_DATASET As Long = &H108       ' dataset key word
Public Const LTK_EXPAND As Long = &H109        ' expand key word
Public Const LTK_EXPAND_LIMBO As Long = &H10A  ' expand Limbo key word
Public Const LTK_BETWEEN As Long = &H10B       ' between key word
Public Const LTK_MKEEP As Long = &H10C         ' miss_keep key word
Public Const LTK_MKEEP_LIMBO As Long = &H10D   ' miss_keep Limbo key word
Public Const LTK_MCLEAR As Long = &H10E        ' miss_clear key word
Public Const LTK_IDENT As Long = &H10F         ' an identifier
Public Const LTK_STRING As Long = &H110        ' a quoted string
Public Const LTK_INTEGER As Long = &H111       ' a number

' Maximum length for any TEXT value
Private Const YYLMAX As Long = 255

' start conditions
Private Const SC_INITIAL As Long = 0
Private Const SC_COMMENT As Long = 1
Private Const SC_SLCOMMENT As Long = 2
Private Const SC_QUOTE As Long = 3

Private inBuf() As Byte                 ' YY_EXTRA_TYPE: the input buffer
Private inLen As Long
Private inPos As Long
Private yy_init As Boolean
Private yylineno As Long
Private yytext As String
Private yy_start As Long

' lexerinput_t, LEXINPUT_BUFFER: buffer(0 .. size - 1)
Public Sub lev_set_extra(buffer() As Byte, ByVal size As Long)
    Dim i As Long
    inLen = size
    If size > 0 Then
        ReDim inBuf(0 To size - 1)
        For i = 0 To size - 1
            inBuf(i) = buffer(i)
        Next
    End If
    inPos = 0
    yy_init = False
End Sub

Public Function lev_get_lineno() As Long
    If Not yy_init Then lev_get_lineno = 1 Else lev_get_lineno = yylineno
End Function

Public Function lev_get_text() As String
    lev_get_text = yytext
End Function

' For non-reentrant C scanner only.
Public Function lev_lex_destroy() As Long
    Erase inBuf
    inLen = 0
    inPos = 0
    yy_init = False
    yytext = ""
    lev_lex_destroy = 0
End Function

Private Function IsAlpha(ByVal c As Long) As Boolean
    IsAlpha = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

Private Function IsDigit(ByVal c As Long) As Boolean
    IsDigit = (c >= 48 And c <= 57)
End Function

' [-0-9_a-zA-Z]
Private Function IsIdentChar(ByVal c As Long) As Boolean
    IsIdentChar = IsAlpha(c) Or IsDigit(c) Or c = 45 Or c = 95
End Function

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

' sstrcpy(aText, yytext); pLevToken = aText;
Private Sub SetToken()
    If Len(yytext) > YYLMAX - 1 Then pLevToken = Left$(yytext, YYLMAX - 1) Else pLevToken = yytext
End Sub

Private Function KeywordToken(ByVal s As String) As Long
    Select Case s
    Case "level": KeywordToken = LTK_LEVEL
    Case "players": KeywordToken = LTK_PLAYERS
    Case "type": KeywordToken = LTK_TYPE
    Case "data": KeywordToken = LTK_DATA
    Case "game": KeywordToken = LTK_GAME
    Case "campaign": KeywordToken = LTK_CAMPAIGN
    Case "camstart": KeywordToken = LTK_CAMSTART
    Case "camchange": KeywordToken = LTK_CAMCHANGE
    Case "dataset": KeywordToken = LTK_DATASET
    Case "expand": KeywordToken = LTK_EXPAND
    Case "expand_limbo": KeywordToken = LTK_EXPAND_LIMBO
    Case "between": KeywordToken = LTK_BETWEEN
    Case "miss_keep": KeywordToken = LTK_MKEEP
    Case "miss_keep_limbo": KeywordToken = LTK_MKEEP_LIMBO
    Case "miss_clear": KeywordToken = LTK_MCLEAR
    End Select
End Function

Public Function lev_lex() As Long
    Dim c As Long, n As Long, k As Long
    If Not yy_init Then
        yylineno = 1
        yy_start = SC_INITIAL
        yy_init = True
    End If
    Do
        If inPos >= inLen Then
            ' <<EOF>>: yyterminate()
            yytext = ""
            lev_lex = 0
            Exit Function
        End If
        c = inBuf(inPos)
        Select Case yy_start
        Case SC_INITIAL
            If IsAlpha(c) Then
                ' Keywords / Match text values: [a-zA-Z][-0-9_a-zA-Z]*
                n = 1
                Do While inPos + n < inLen
                    If Not IsIdentChar(inBuf(inPos + n)) Then Exit Do
                    n = n + 1
                Loop
                Match n
                k = KeywordToken(yytext)
                If k <> 0 Then
                    lev_lex = k
                Else
                    SetToken
                    lev_lex = LTK_IDENT
                End If
                Exit Function
            ElseIf IsDigit(c) Or (c = 45 And IsDigit(PeekByte(1))) Then
                ' Match integer numbers: -?[0-9]+
                n = 1
                Do While inPos + n < inLen
                    If Not IsDigit(inBuf(inPos + n)) Then Exit Do
                    n = n + 1
                Loop
                Match n
                levVal = atoi(yytext) And &HFF&       ' (LEVEL_TYPE)atol(yytext): an uint8_t
                lev_lex = LTK_INTEGER
                Exit Function
            ElseIf c = 34 Then
                ' Match quoted text
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
                If c >= 128 Then lev_lex = c - 256 Else lev_lex = c
                Exit Function
            End If
        Case SC_QUOTE
            If c = 34 Then
                Match 1
                yy_start = SC_INITIAL
            Else
                ' [^\"]* (a lone newline is the earlier rule's)
                n = 1
                Do While inPos + n < inLen
                    If inBuf(inPos + n) = 34 Then Exit Do
                    n = n + 1
                Loop
                If n = 1 And c = 10 Then
                    Match 1
                    lev_error "Unexpected end of line in string"
                Else
                    Match n
                    SetToken
                    lev_lex = LTK_STRING
                    Exit Function
                End If
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
