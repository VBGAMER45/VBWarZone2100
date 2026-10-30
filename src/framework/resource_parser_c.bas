Attribute VB_Name = "resource_parser_c"
' lib/framework/resource_parser.ypp: the bison parser for resource (*.wrf)
' files (prefix res_), written out by hand. The grammar
'     res_file:  res_line | res_file res_line
'     res_line:  dir_line | file_line
'     dir_line:  DIRECTORY QTEXT_T
'     file_line: FILETOKEN TEXT_T QTEXT_T
' is read token by token, as the LALR parser reads it: each line's action
' runs before the next token is read; an unexpected token (or an empty
' file) is bison's "syntax error" through yyerror (res_error), and the
' parser returns 1, as it does for YYABORT; a token <= 0 is the end.
' aResDir and aCurrResDir are frameresource_c's.
Option Explicit

Private Const PATH_MAX As Long = 260

' yyerror (res_error with the prefix)
Public Sub res_error(ByVal msg As String)
    debug_ 46, LOG_ERROR, "res_error", "RES file parse error:\n%s at line %d\nText: '%s'\n", msg, res_get_lineno(), res_get_text()
End Sub

' dir_line: DIRECTORY QTEXT_T
Private Sub dir_line(ByVal s2 As String)
    Dim lenDir As Long
    ' set a new input directory
    debug_ 90, LOG_NEVER, "res_parse", "directory: %s", s2
    If Left$(s2, 2) = "/:" Then
        ' the new dir is rooted
        aCurrResDir = Left$(s2, PATH_MAX - 1)
    Else
        aCurrResDir = aResDir
        aCurrResDir = Left$(aCurrResDir & s2, PATH_MAX - 1)
    End If
    If Len(s2) > 0 Then
        If Not (WZ_PHYSFS_isDirectory(aCurrResDir) <> 0) Then ASSERT_FAILURE "WZ_PHYSFS_isDirectory(aCurrResDir)", "lib/framework/resource_parser.ypp:103", "res_parse", "%s is not a directory!", aCurrResDir
        ' Add a trailing '/'
        aCurrResDir = aCurrResDir & "/"
        debug_ 108, LOG_NEVER, "res_parse", "Current resource directory: %s", aCurrResDir
    End If
End Sub

' file_line: FILETOKEN TEXT_T QTEXT_T (false: YYABORT)
Private Function file_line(ByVal s2 As String, ByVal s3 As String) As Boolean
    Dim success As Boolean
    ' load a data file
    debug_ 119, LOG_NEVER, "res_parse", "file: %s %s", s2, s3
    success = resLoadFile(s2, s3)

    file_line = success
End Function

' 0: accepted; 1: a syntax error, or an action's YYABORT
Public Function res_parse() As Long
    Dim tok As Long, lines As Long, s2 As String, s3 As String
    Do
        tok = res_lex()
        If tok <= 0 Then
            ' end of input: res_file needs a res_line
            If lines > 0 Then
                res_parse = 0
            Else
                res_error "syntax error"
                res_parse = 1
            End If
            Exit Function
        End If
        Select Case tok
        Case DIRECTORY
            If res_lex() <> QTEXT_T Then GoTo syntax_error
            s2 = res_lval_sval
            dir_line s2
        Case FILETOKEN
            If res_lex() <> TEXT_T Then GoTo syntax_error
            s2 = res_lval_sval
            If res_lex() <> QTEXT_T Then GoTo syntax_error
            s3 = res_lval_sval
            If Not file_line(s2, s3) Then
                res_parse = 1                       ' YYABORT
                Exit Function
            End If
        Case Else
            GoTo syntax_error
        End Select
        lines = lines + 1
    Loop
syntax_error:
    res_error "syntax error"
    res_parse = 1
End Function
