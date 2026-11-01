Attribute VB_Name = "map_levparse_c"
' lib/wzmaplib/src/map_levparse.cpp, map_levparse.h: parsing .lev files (levParseBasic), the classic dataset and
' level type names. levParseBasic's two std::function handlers are its result, in order: the events
' evKind(0 .. evCount - 1) (LEVEV_COMMAND: evA the command, evB its argument; LEVEV_COMMENT: evA the comment line),
' which the caller handles one after another as the C++ handlers run. The C++ strings hold the file's bytes: a
' token is built a byte at a time (one character per byte) and decoded as UTF-8 when it is handed over.
Option Explicit

' enum class SUPPORTED_LEVEL_TYPES : int
Public Const SUPPORTED_LEVEL_TYPES_CAMPAIGN As Long = 12
Public Const SUPPORTED_LEVEL_TYPES_SKIRMISH As Long = 14
Public Const SUPPORTED_LEVEL_TYPES_MULTI_SKIRMISH2 As Long = 18 ' all these old SKIRMISH entries in .lev files should be mapped to SKIRMISH
Public Const SUPPORTED_LEVEL_TYPES_MULTI_SKIRMISH3 As Long = 19
Public Const SUPPORTED_LEVEL_TYPES_MULTI_SKIRMISH4 As Long = 20

' (VB) levParseBasic's events: handleCommand(command, arg), handleCommentLine(commentLine)
Public Const LEVEV_COMMAND As Long = 1
Public Const LEVEV_COMMENT As Long = 2

' enum class ParsingMode
Private Const ParsingMode_Normal As Long = 0
Private Const ParsingMode_QuotedText As Long = 1
Private Const ParsingMode_SingleLineComment As Long = 2
Private Const ParsingMode_Comment As Long = 3

' the parse's state (the C++ lambdas' captures)
Private mode As Long
Private currentCommandToken(0 To 2) As String
Private currentCommandTokenIdx As Long
Private currentCommentLine As String
Private commentStartedAtBeginningOfLine As Boolean
Private line_ As Long
Private lastModeChangeStartPos As Long
Private idx As Long
Private pLogger As LoggingProtocol
Private ek() As Long, ea() As String, eb() As String, en As Long

' (VB) a token's bytes (one character each) as the UTF-8 text they are
Private Function tokenText(ByVal s As String) As String
    Dim b() As Byte, i As Long
    If Len(s) = 0 Then Exit Function
    ReDim b(0 To Len(s) - 1)
    For i = 1 To Len(s)
        b(i - 1) = AscW(Mid$(s, i, 1)) And &HFF&
    Next
    tokenText = utf8_decode(b, 0, Len(s))
End Function

Private Sub addEvent(ByVal kind As Long, ByVal a As String, ByVal b As String)
    ReDim Preserve ek(0 To en)
    ReDim Preserve ea(0 To en)
    ReDim Preserve eb(0 To en)
    ek(en) = kind
    ea(en) = tokenText(a)
    eb(en) = tokenText(b)
    en = en + 1
End Sub

Private Sub addCharToCommandPart(ByVal c As Long)
    If currentCommandTokenIdx <= UBound(currentCommandToken) Then
        currentCommandToken(currentCommandTokenIdx) = currentCommandToken(currentCommandTokenIdx) & ChrW$(c)
    End If
End Sub

Private Sub handleLine()
    Dim k As Long
    If mode <> ParsingMode_Comment Then
        mode = ParsingMode_Normal
        commentStartedAtBeginningOfLine = False
        lastModeChangeStartPos = idx + 1
    Else
        commentStartedAtBeginningOfLine = True
    End If
    If Len(currentCommandToken(0)) > 0 Then
        If Len(currentCommandToken(2)) > 0 Then
            ' LEV FORMAT ERROR: More than two tokens on a line
            ' (Expected: <command> <argument>)
            printLog_ pLogger, 92, LogLevel_Error, "levParseBasic", "LEV File Parse Error: `Syntax Error` at line %zu: at least one unexpected token: %s", line_, tokenText(currentCommandToken(2))
        End If
        addEvent LEVEV_COMMAND, currentCommandToken(0), currentCommandToken(1)
        For k = 0 To 2
            currentCommandToken(k) = ""
        Next
        currentCommandTokenIdx = 0
    End If
    If Len(currentCommentLine) > 0 Then
        addEvent LEVEV_COMMENT, currentCommentLine, ""
        currentCommentLine = ""
    End If
End Sub

' data(0 .. dataLen - 1); evKind / evA / evB (0 .. evCount - 1): the handlers' calls in order
Public Function levParseBasic(data() As Byte, ByVal dataLen As Long, ByVal mapIO As IOProvider, evKind() As Long, evA() As String, _
        evB() As String, evCount As Long, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim c As Long, nextChar As Long, startPosOfLine As Long, k As Long
    evCount = 0
    If dataLen = 0 Then
        levParseBasic = False
        Exit Function
    End If

    Set pLogger = pCustomLogger
    mode = ParsingMode_Normal
    For k = 0 To 2
        currentCommandToken(k) = ""
    Next
    currentCommandTokenIdx = 0
    currentCommentLine = ""
    commentStartedAtBeginningOfLine = False
    line_ = 0
    lastModeChangeStartPos = 0
    startPosOfLine = 0
    en = 0
    Erase ek: Erase ea: Erase eb

    idx = 0
    Do While idx < dataLen
        c = data(idx)
        ' lookAhead(1): nextChar.value_or(0)
        If idx + 1 >= dataLen Then nextChar = 0 Else nextChar = data(idx + 1)
        If c = 10 Or c = 13 Or c = 0 Then
            ' handle newline (or NULL)
            handleLine
            line_ = line_ + 1
            startPosOfLine = idx + 1
        ElseIf mode = ParsingMode_QuotedText Then
            If c = 34 Then
                mode = ParsingMode_Normal
                lastModeChangeStartPos = idx + 1
            Else
                addCharToCommandPart c
            End If
        ElseIf c = 47 And (nextChar = 47 Or nextChar = 42) And (mode <> ParsingMode_Comment And mode <> ParsingMode_SingleLineComment) Then
            If nextChar = 47 Then
                mode = ParsingMode_SingleLineComment
            Else
                mode = ParsingMode_Comment
            End If
            commentStartedAtBeginningOfLine = (idx = startPosOfLine)
            idx = idx + 1
            lastModeChangeStartPos = idx + 1
        ElseIf mode = ParsingMode_Comment And c = 42 And nextChar = 47 Then
            mode = ParsingMode_Normal
            If Len(currentCommentLine) > 0 Then
                addEvent LEVEV_COMMENT, currentCommentLine, ""
                currentCommentLine = ""
            End If
            idx = idx + 1
            lastModeChangeStartPos = idx + 1
        ElseIf mode = ParsingMode_Normal And c = 34 Then
            mode = ParsingMode_QuotedText
            lastModeChangeStartPos = idx + 1
        Else
            If mode = ParsingMode_Normal And (c = 32 Or c = 9 Or c = 12) Then
                ' whitespace is treated as a separator for command and arg - if a command has already been read
                If currentCommandTokenIdx <= UBound(currentCommandToken) Then
                    If Len(currentCommandToken(currentCommandTokenIdx)) > 0 Then
                        currentCommandTokenIdx = currentCommandTokenIdx + 1
                    End If
                End If
                ' otherwise, just ignore it - must be whitespace prior to a command
            Else
                ' append the character to the appropriate type of token (if needed)
                Select Case mode
                    Case ParsingMode_Normal, ParsingMode_QuotedText
                        addCharToCommandPart c
                    Case ParsingMode_SingleLineComment, ParsingMode_Comment
                        If commentStartedAtBeginningOfLine Then
                            currentCommentLine = currentCommentLine & ChrW$(c)
                        End If
                End Select
            End If
        End If
        idx = idx + 1
    Loop

    ' handle last line (if present)
    handleLine

    evCount = en
    If en > 0 Then
        evKind = ek
        evA = ea
        evB = eb
    End If
    Set pLogger = Nothing
    levParseBasic = True
End Function

' sscanf(s, "MULTI_CAM_%u", ...) / "MULTI_T%u_C%u": an unsigned decimal after the prefix (leading spaces
' skipped, as %u does); the text after it is returned in rest
Private Function scanUnsigned(ByVal s As String, ByVal prefix As String, value As Long, rest As String) As Boolean
    Dim i As Long, digits As String
    If Left$(s, Len(prefix)) <> prefix Then Exit Function
    i = Len(prefix) + 1
    Do While Mid$(s, i, 1) = " " Or Mid$(s, i, 1) = vbTab
        i = i + 1
    Loop
    Do While i <= Len(s) And Mid$(s, i, 1) >= "0" And Mid$(s, i, 1) <= "9"
        digits = digits & Mid$(s, i, 1)
        i = i + 1
    Loop
    If Len(digits) = 0 Then Exit Function
    If Len(digits) > 9 Then value = -1 Else value = CLng(digits)     ' (a larger value maps to no tileset)
    rest = Mid$(s, i)
    scanUnsigned = True
End Function

' optional<MAP_TILESET>: True with the tileset
Public Function convertLevMapTilesetType(ByVal dataset As String, tileset As MAP_TILESET) As Boolean
    Dim t As Long, c As Long, parsedDataset As Boolean, rest As String
    ' MULTI_CAM_1, MULTI_CAM_2, MULTI_CAM_3
    parsedDataset = scanUnsigned(dataset, "MULTI_CAM_", c, rest)
    If Not parsedDataset Then
        ' Many older maps also specify > 1 tech level variants. Examples: MULTI_T2_C2, MULTI_T3_C2, etc
        ' Where T# is the tech level
        ' And C# maps to the campaign/tileset (which is all we care about)
        c = 0
        If scanUnsigned(dataset, "MULTI_T", t, rest) Then
            parsedDataset = scanUnsigned(rest, "_C", c, rest)
            If Not parsedDataset Then c = 0
        End If
    End If

    Select Case c
    Case 1
        tileset = MAP_TILESET_ARIZONA
        convertLevMapTilesetType = True
    Case 2
        tileset = MAP_TILESET_URBAN
        convertLevMapTilesetType = True
    Case 3
        tileset = MAP_TILESET_ROCKIES
        convertLevMapTilesetType = True
    End Select
End Function
