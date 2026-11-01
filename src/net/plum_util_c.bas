Attribute VB_Name = "plum_util_c"
' libplum: Copyright (c) 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/util.c, util.h. A const char * into a string is its 1-based position in that String (0: NULL);
' a char buffer of size bytes is a ByRef String that gets at most size - 1 characters, and the functions return
' snprintf's count as C's do. xml_find_matching_child returns the text from the match on, of the copy it searched
' (C returns a pointer into that copy after freeing it; its callers read it as it was).
Option Explicit

Private Const PATTERN_BUFFER_SIZE As Long = 256

' (VB) toupper() of each ASCII letter, as C's locale does it
Private Function toupper_ascii(ByVal s As String) As String
    Dim b() As Byte, i As Long
    If Len(s) = 0 Then Exit Function
    b = s
    For i = 0 To UBound(b) - 1 Step 2
        If b(i + 1) = 0 Then
            If b(i) >= 97 And b(i) <= 122 Then b(i) = b(i) - 32
        End If
    Next
    toupper_ascii = b
End Function

' Case-insensitive equivalent to strstr(): static char *my_stristr(const char *haystack, const char *needle),
' searching from position start
' (VB) the first match its loops find, by InStr on the upper-cased texts
Private Function my_stristr(ByVal haystack As String, ByVal start As Long, ByVal needle As String) As Long
    If start > Len(haystack) Then Exit Function
    my_stristr = InStr(start, toupper_ascii(haystack), toupper_ascii(needle), vbBinaryCompare)
End Function

' static char *xml_remove_attributes(const char *str)
Private Function xml_remove_attributes(ByVal str_ As String) As String
    Dim p1 As Long, c As String, inside_tag As Boolean, skip_chars As Boolean, out As String, n As Long
    n = Len(str_)
    out = Space$(n)
    Dim o As Long
    inside_tag = False
    skip_chars = False
    For p1 = 1 To n
        c = Mid$(str_, p1, 1)
        If c = "<" Then
            inside_tag = True
        ElseIf c = " " And inside_tag Then
            skip_chars = True
        ElseIf c = ">" Then
            inside_tag = False: skip_chars = False
        End If
        If Not skip_chars Then
            o = o + 1
            Mid$(out, o, 1) = c
        End If
    Next
    xml_remove_attributes = Left$(out, o)
End Function

' const char *my_find(const char *str, const char *before, bool case_insensitive): from position start
Public Function my_find(ByVal str_ As String, ByVal start As Long, ByVal before As String, ByVal case_insensitive As Boolean) As Long
    If case_insensitive Then
        my_find = my_stristr(str_, start, before)
    Else
        If start > Len(str_) Then Exit Function
        my_find = InStr(start, str_, before, vbBinaryCompare)   ' strstr(str, before)
    End If
End Function

' (VB) C's isspace
Private Function isspace_(ByVal c As String) As Boolean
    Select Case AscW(c)
    Case 32, 9 To 13: isspace_ = True
    End Select
End Function

' int string_extract(const char *str, const char *before, const char *after, char *buffer, size_t size,
'                    bool case_insensitive)
Public Function string_extract(ByVal str_ As String, ByVal before As String, ByVal after As String, buffer As String, _
        ByVal size As Long, ByVal case_insensitive As Boolean) As Long
    Dim begin_ As Long, end_ As Long, len_ As Long
    begin_ = my_find(str_, 1, before, case_insensitive)
    If begin_ = 0 Then
        string_extract = -1
        Exit Function
    End If

    begin_ = begin_ + Len(before)

    end_ = my_find(str_, begin_, after, case_insensitive)
    If end_ = 0 Then end_ = Len(str_) + 1

    Do While begin_ <> end_
        If Not isspace_(Mid$(str_, begin_, 1)) Then Exit Do
        begin_ = begin_ + 1
    Loop

    Do While begin_ <> end_
        If Not isspace_(Mid$(str_, end_ - 1, 1)) Then Exit Do
        end_ = end_ - 1
    Loop

    len_ = end_ - begin_
    ' snprintf(buffer, size, "%.*s", len, begin)
    If len_ > size - 1 Then buffer = Mid$(str_, begin_, size - 1) Else buffer = Mid$(str_, begin_, len_)
    string_extract = len_
End Function

' int header_extract(const char *str, const char *name, char *buffer, size_t size)
Public Function header_extract(ByVal str_ As String, ByVal name_ As String, buffer As String, ByVal size As Long) As Long
    Dim before As String, after As String
    If name_ = "" Then
        header_extract = -1
        Exit Function
    End If

    before = vbCrLf & name_ & ":"               ' snprintf(before, PATTERN_BUFFER_SIZE, "\r\n%s:", name)
    If Len(before) >= PATTERN_BUFFER_SIZE Then
        header_extract = -1
        Exit Function
    End If

    after = vbCrLf

    header_extract = string_extract(str_, before, after, buffer, size, True) ' case-insensitive
End Function

' int xml_extract(const char *str, const char *tag, char *buffer, size_t size)
Public Function xml_extract(ByVal str_ As String, ByVal tag As String, buffer As String, ByVal size As Long) As Long
    Dim before As String, after As String, str2 As String
    If tag = "" Then
        xml_extract = -1
        Exit Function
    End If

    before = "<" & tag & ">"
    If Len(before) >= PATTERN_BUFFER_SIZE Then
        xml_extract = -1
        Exit Function
    End If

    after = "</" & tag & ">"
    If Len(after) >= PATTERN_BUFFER_SIZE Then
        xml_extract = -1
        Exit Function
    End If

    str2 = xml_remove_attributes(str_)
    xml_extract = string_extract(str2, before, after, buffer, size, True) ' case-insensitive
End Function

' const char *xml_find_matching_child(const char *str, const char *tag, const char *child_tag,
'                                     const char *child_value): the text from the match on ("" for NULL), found
' says whether it matched
Public Function xml_find_matching_child(ByVal str_ As String, ByVal tag As String, ByVal child_tag As String, _
        ByVal child_value As String, found As Boolean) As String
    Dim before As String, after As String, child_before As String, child_after As String, str2 As String
    Dim pos As Long, end_ As Long, child_pos As Long, child_end As Long
    found = False
    If tag = "" Or child_tag = "" Then Exit Function

    before = "<" & tag & ">"
    If Len(before) >= PATTERN_BUFFER_SIZE Then Exit Function

    after = "</" & tag & ">"
    If Len(after) >= PATTERN_BUFFER_SIZE Then Exit Function

    child_before = "<" & child_tag & ">"
    If Len(child_before) >= PATTERN_BUFFER_SIZE Then Exit Function

    child_after = "</" & child_tag & ">"
    If Len(child_after) >= PATTERN_BUFFER_SIZE Then Exit Function

    str2 = xml_remove_attributes(str_)
    pos = 1
    Do
        pos = my_find(str2, pos, before, True)
        If pos = 0 Then Exit Do
        pos = pos + Len(before)

        end_ = my_find(str2, pos, after, True)
        If end_ = 0 Then end_ = Len(str2) + 1

        child_pos = pos
        Do
            child_pos = my_find(str2, child_pos, child_before, True)
            If child_pos = 0 Then Exit Do
            If child_pos >= end_ Then Exit Do
            child_pos = child_pos + Len(child_before)

            child_end = my_find(str2, child_pos, child_after, True)
            If child_end = 0 Then child_end = end_

            ' strncmp(child_value, child_pos, child_end - child_pos) == 0
            If strncmp_eq(child_value, Mid$(str2, child_pos), child_end - child_pos) Then
                found = True
                xml_find_matching_child = Mid$(str2, pos)
                Exit Function
            End If
        Loop
    Loop

End Function

' (VB) strncmp(a, b, n) == 0 for C strings (the comparison stops at a NUL, the end of a String here)
Private Function strncmp_eq(ByVal a As String, ByVal b As String, ByVal n As Long) As Boolean
    Dim i As Long, ca As Long, cb As Long
    For i = 1 To n
        If i <= Len(a) Then ca = AscW(Mid$(a, i, 1)) Else ca = 0
        If i <= Len(b) Then cb = AscW(Mid$(b, i, 1)) Else cb = 0
        If ca <> cb Then Exit Function
        If ca = 0 Then Exit For
    Next
    strncmp_eq = True
End Function
