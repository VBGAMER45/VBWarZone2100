Attribute VB_Name = "piepalette_c"
' lib/ivis_opengl/piepalette.cpp, lib/ivis_opengl/piepalette.h: the palette
' (data/base/palette.txt). A PIELIGHT is a Long, its rgba() (r in the low byte,
' as the union's bytes lie in memory). WZCOL_X (psPalette[n]) is
' psPalette(WZCOL_X), the index being GfxApi's WZCOL enum (WZCOL_MAX = 101).
' The header's static inlines pal_Colour, pal_SetBrightness and pal_RGBA are
' here for WzGfx (the exe has its own copies).
Option Explicit

Public psPalette(0 To WZCOL_MAX - 1) As Long ' PIELIGHT psPalette[WZCOL_MAX];

Public Sub pal_Init()
    Dim pFileData() As Byte, fileSize As Long, chars As String, ptr As Long, lenLeft As Long
    Dim i As Long, r As Long, g As Long, b As Long, a As Long, len_ As Long, result As Long

    ' Read these from file so that mod-makers can change them
    If loadFile("palette.txt", pFileData, fileSize) Then
        ' (VB) the bytes as characters, one each, so that offsets are the C++ pointer's
        chars = String$(fileSize, vbNullChar)
        For i = 0 To fileSize - 1
            Mid$(chars, i + 1, 1) = ChrW$(pFileData(i))
        Next
        ptr = 0
        lenLeft = fileSize
        For i = 0 To WZCOL_MAX - 1
            If lenLeft <= 0 Then
                ' No more data to read - palette.txt is likely for an older version of WZ
                debug_ 43, LOG_FATAL, "pal_Init", "palette.txt is missing required value(s); (missing lines %d - %d)", i, (WZCOL_MAX - 1)
                Exit Sub
            End If

            len_ = 0                    ' int len; (C++ leaves it unset when %n is not reached)

            result = sscanf(Mid$(chars, ptr + 1), "%x, %x, %x, %x %*[^" & vbLf & "]" & vbLf & "%n", r, g, b, a, len_)

            ptr = ptr + len_
            lenLeft = lenLeft - len_
            If Not (lenLeft >= 0) Then ASSERT_FAILURE "lenLeft >= 0", "lib/ivis_opengl/piepalette.cpp:54", "pal_Init", "Buffer overrun reading palette data"

            If result <> 4 Then
                debug_ 58, LOG_ERROR, "pal_Init", "Truncated / invalid line (%d) in palette.txt only has %d valid values", (i + 1), result
                GoTo continue_
            End If
            psPalette(i) = pal_RGBA(r, g, b, a)
continue_:
        Next
        Erase pFileData
    Else
        debug_ 70, LOG_FATAL, "pal_Init", "Failed to load palette.txt"
    End If
End Sub

Public Sub pal_ShutDown()
    ' placeholder
End Sub

Public Function pal_GetTeamColour(ByVal team As Long) As Long
    Dim tcolour As Long

    ' set correct team colour based on team
    Select Case team
    Case 0
        tcolour = psPalette(WZCOL_TEAM1) 'green
    Case 1
        tcolour = psPalette(WZCOL_TEAM2) 'orange
    Case 2
        tcolour = psPalette(WZCOL_TEAM3) 'gray
    Case 3
        tcolour = psPalette(WZCOL_TEAM4) 'black
    Case 4
        tcolour = psPalette(WZCOL_TEAM5) 'red
    Case 5
        tcolour = psPalette(WZCOL_TEAM6) 'blue
    Case 6
        tcolour = psPalette(WZCOL_TEAM7) 'purple
    Case 7
        tcolour = psPalette(WZCOL_TEAM8) 'teal
    Case 8: pal_GetTeamColour = psPalette(WZCOL_TEAM9): Exit Function
    Case 9: pal_GetTeamColour = psPalette(WZCOL_TEAM10): Exit Function
    Case 10: pal_GetTeamColour = psPalette(WZCOL_TEAM11): Exit Function
    Case 11: pal_GetTeamColour = psPalette(WZCOL_TEAM12): Exit Function
    Case 12: pal_GetTeamColour = psPalette(WZCOL_TEAM13): Exit Function
    Case 13: pal_GetTeamColour = psPalette(WZCOL_TEAM14): Exit Function
    Case 14: pal_GetTeamColour = psPalette(WZCOL_TEAM15): Exit Function
    Case 15: pal_GetTeamColour = psPalette(WZCOL_TEAM16): Exit Function
    Case Else
        ' STATIC_ASSERT(MAX_PLAYERS <= 16);
        ASSERT_FAILURE "false", "lib/ivis_opengl/piepalette.cpp:120", "pal_GetTeamColour", "Attempting to get colour for non-existing team %u", team
        tcolour = psPalette(WZCOL_WHITE) 'default is white
    End Select

    pal_GetTeamColour = tcolour
End Function

' piepalette.h's static inlines (PIELIGHT { r, g, b, a } as its rgba())
Public Function pal_Colour(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    pal_Colour = pal_RGBA(r, g, b, &HFF&)
End Function

Public Function pal_SetBrightness(ByVal brightness As Long) As Long
    pal_SetBrightness = pal_RGBA(brightness, brightness, brightness, &HFF&)
End Function

Public Function pal_RGBA(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal a As Long) As Long
    a = a And &HFF&
    ' a << 24 without an overflow: the high byte is the sign
    If a >= &H80& Then
        pal_RGBA = (r And &HFF&) Or ((g And &HFF&) * &H100&) Or ((b And &HFF&) * &H10000) Or ((a - &H100&) * &H1000000)
    Else
        pal_RGBA = (r And &HFF&) Or ((g And &HFF&) * &H100&) Or ((b And &HFF&) * &H10000) Or (a * &H1000000)
    End If
End Function
