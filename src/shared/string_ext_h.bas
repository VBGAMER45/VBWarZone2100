Attribute VB_Name = "string_ext_h"
' lib/framework/string_ext.h's inline functions (compiled into every
' project, as the header is). Ported: strEndsWith, strJoin.
Option Explicit

Public Function strEndsWith(ByVal str_ As String, ByVal suffix As String) As Boolean
    strEndsWith = (Len(str_) >= Len(suffix)) And (Right$(str_, Len(suffix)) = suffix)
End Function

Public Function strJoin(strs As StringVector, ByVal sep As String) As String
    Dim str_ As String, first As Boolean, i As Long
    first = True
    For i = 0 To strs.size - 1
        If Not first Then
            str_ = str_ & sep
        End If
        str_ = str_ & strs.items(i)
        first = False
    Next
    strJoin = str_
End Function
