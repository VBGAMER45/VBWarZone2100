Attribute VB_Name = "strres_parser_c"
' lib/framework/strres_parser.ypp: the bison parser for string resource
' files (prefix strres_), written out by hand. The grammar
'     file:   line | file line
'     line:   TEXT_T string
'     string: QTEXT_T | '_' '(' QTEXT_T ')'
' is read token by token, as the LALR parser reads it: each line's action
' runs before the next token is read; an unexpected token (or an empty
' file) is bison's "syntax error" through yyerror (strres_error), and the
' parser returns 1, as it does for YYABORT; a token <= 0 is the end.
' gettext is not ported: gettext(s) is s.
Option Explicit

' yyerror (strres_error with the prefix)
Public Sub strres_error(ByVal msg As String)
    debug_ 37, LOG_ERROR, "strres_error", "STRRES file parse error:\n%s at line %d\nText: '%s'", msg, strres_get_lineno(), strres_get_text()
End Sub

' 0: accepted; 1: a syntax error, or the action's YYABORT
Public Function strres_parse(ByVal psStrRes As STR_RES) As Long
    Dim tok As Long, lines As Long, s1 As String, s2 As String, s3 As String
    Dim success As Boolean
    Do
        tok = strres_lex()
        If tok <= 0 Then
            ' end of input: file needs a line
            If lines > 0 Then
                strres_parse = 0
            Else
                strres_error "syntax error"
                strres_parse = 1
            End If
            Exit Function
        End If
        ' line: TEXT_T string
        If tok <> TEXT_T Then GoTo syntax_error
        s1 = strres_lval_sval
        tok = strres_lex()
        If tok = QTEXT_T Then
            ' string: QTEXT_T
            s2 = strres_lval_sval
        ElseIf tok = 95 Then
            ' string: '_' '(' QTEXT_T ')'
            If strres_lex() <> 40 Then GoTo syntax_error
            If strres_lex() <> QTEXT_T Then GoTo syntax_error
            s3 = strres_lval_sval
            If strres_lex() <> 41 Then GoTo syntax_error
            s2 = s3                     ' strdup(gettext($3))
        Else
            GoTo syntax_error
        End If

        ' Pass the text string to the string manager
        success = strresStoreString(psStrRes, s1, s2)

        If Not success Then
            strres_parse = 1                    ' YYABORT
            Exit Function
        End If
        lines = lines + 1
    Loop
syntax_error:
    strres_error "syntax error"
    strres_parse = 1
End Function
