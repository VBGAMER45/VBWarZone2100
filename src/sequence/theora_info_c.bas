Attribute VB_Name = "theora_info_c"
' ref/theora (libtheora 1.2.0) lib/info.c: th_info and th_comment.
' (VB) th_comment is theora_h's theora_comment (theora_codec_h's header); a comment string holds its bytes
' as characters 0-255, user_comments(0 To comments - 1) (the C's terminating NULL entry is not kept), and a
' NULL vendor is vbNullString. th_comment_query returns the value as a String (vbNullString for NULL).
Option Explicit

' This is more or less the same as strncasecmp, but that doesn't exist everywhere, and this is a fairly
' trivial function, so we include it. Note: We take advantage of the fact that we know _n is less than or
' equal to the length of at least one of the strings.
Private Function oc_tagcompare(ByVal s1 As String, ByVal s2 As String, ByVal n As Long) As Long
    Dim c As Long
    For c = 0 To n - 1
        If oc_vbToupper(oc_vbCharAt(s1, c)) <> oc_vbToupper(oc_vbCharAt(s2, c)) Then oc_tagcompare = 1: Exit Function
    Next
    If oc_vbCharAt(s1, n) <> 61 Then oc_tagcompare = 1 Else oc_tagcompare = 0  ' _s1[c]!='='
End Function

' (VB) the C string's byte at index i (0 past the end: its terminating NUL)
Private Function oc_vbCharAt(s As String, ByVal i As Long) As Long
    If i < Len(s) Then oc_vbCharAt = AscW(Mid$(s, i + 1, 1)) And 255
End Function

' (VB) toupper in the "C" locale
Private Function oc_vbToupper(ByVal c As Long) As Long
    If c >= 97 And c <= 122 Then oc_vbToupper = c - 32 Else oc_vbToupper = c
End Function

Public Sub th_info_init(info As th_info)
    Dim blank As th_info
    info = blank
    info.version_major = TH_VERSION_MAJOR
    info.version_minor = TH_VERSION_MINOR
    info.version_subminor = TH_VERSION_SUB
    info.keyframe_granule_shift = 6
End Sub

Public Sub th_info_clear(info As th_info)
    Dim blank As th_info
    info = blank
End Sub

Public Sub th_comment_init(tc As theora_comment)
    Dim blank As theora_comment
    tc = blank
End Sub

Public Sub th_comment_add(tc As theora_comment, ByVal comment As String)
    Dim comment_len As Long
    ' (VB) user_comments=_ogg_realloc(_tc->user_comments,(_tc->comments+2)*sizeof(*_tc->user_comments))
    If tc.comments = 0 Then
        ReDim tc.user_comments(0 To 1)
        ReDim tc.comment_lengths(0 To 1)
    Else
        ReDim Preserve tc.user_comments(0 To tc.comments + 1)
        ReDim Preserve tc.comment_lengths(0 To tc.comments + 1)
    End If
    comment_len = Len(comment)
    tc.comment_lengths(tc.comments) = comment_len
    tc.user_comments(tc.comments) = comment
    tc.comments = tc.comments + 1
    tc.user_comments(tc.comments) = vbNullString
End Sub

Public Sub th_comment_add_tag(tc As theora_comment, ByVal tag As String, ByVal val_ As String)
    Dim comment As String
    comment = tag & "=" & val_
    th_comment_add tc, comment
End Sub

Public Function th_comment_query(tc As theora_comment, ByVal tag As String, ByVal count As Long) As String
    Dim i As Long
    Dim found As Long
    Dim tag_len As Long
    tag_len = Len(tag)
    found = 0
    For i = 0 To tc.comments - 1
        If oc_tagcompare(tc.user_comments(i), tag, tag_len) = 0 Then
            ' We return a pointer to the data, not a copy.
            If count = found Then
                th_comment_query = Mid$(tc.user_comments(i), tag_len + 2)
                Exit Function
            End If
            found = found + 1
        End If
    Next
    ' Didn't find anything.
    th_comment_query = vbNullString
End Function

Public Function th_comment_query_count(tc As theora_comment, ByVal tag As String) As Long
    Dim i As Long
    Dim tag_len As Long
    Dim count As Long
    tag_len = Len(tag)
    count = 0
    For i = 0 To tc.comments - 1
        If oc_tagcompare(tc.user_comments(i), tag, tag_len) = 0 Then count = count + 1
    Next
    th_comment_query_count = count
End Function

Public Sub th_comment_clear(tc As theora_comment)
    Dim blank As theora_comment
    ' for(i=0;i<_tc->comments;i++)_ogg_free(_tc->user_comments[i]); _ogg_free(...): (VB) the strings and
    ' arrays go with the struct's contents
    tc = blank
End Sub
