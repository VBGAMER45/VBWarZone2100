Attribute VB_Name = "vector_vb"
' std::vector<std::string> for the port (no C++ counterpart; compiled into
' every project): a dynamic array plus a count (docs/CONVENTIONS.md). The
' items are items(0 .. size - 1).
Option Explicit

Public Type StringVector
    items() As String
    size As Long
End Type

' C: v.push_back(s)
Public Sub StringVector_push_back(v As StringVector, ByVal s As String)
    If v.size = 0 Then
        ReDim v.items(0 To 7)
    ElseIf v.size > UBound(v.items) Then
        ReDim Preserve v.items(0 To 2 * v.size - 1)
    End If
    v.items(v.size) = s
    v.size = v.size + 1
End Sub

' C: v.erase(v.begin() + i)
Public Sub StringVector_erase(v As StringVector, ByVal i As Long)
    Dim j As Long
    For j = i To v.size - 2
        v.items(j) = v.items(j + 1)
    Next
    v.size = v.size - 1
    v.items(v.size) = vbNullString
End Sub

' C: std::find(v.begin(), v.end(), s) != v.end()
Public Function StringVector_contains(v As StringVector, ByVal s As String) As Boolean
    Dim i As Long
    For i = 0 To v.size - 1
        If v.items(i) = s Then
            StringVector_contains = True
            Exit Function
        End If
    Next
End Function

' C: std::sort(v.begin(), v.end()) (std::string's operator<: binary order)
Public Sub StringVector_sort(v As StringVector)
    Dim i As Long, j As Long, s As String
    For i = 1 To v.size - 1
        s = v.items(i)
        j = i - 1
        Do While j >= 0
            If StrComp(v.items(j), s, vbBinaryCompare) <= 0 Then Exit Do
            v.items(j + 1) = v.items(j)
            j = j - 1
        Loop
        v.items(j + 1) = s
    Next
End Sub

' C: v.erase(std::unique(v.begin(), v.end()), v.end())
Public Sub StringVector_unique(v As StringVector)
    Dim i As Long, n As Long
    If v.size = 0 Then Exit Sub
    n = 1
    For i = 1 To v.size - 1
        If v.items(i) <> v.items(n - 1) Then
            v.items(n) = v.items(i)
            n = n + 1
        End If
    Next
    For i = n To v.size - 1
        v.items(i) = vbNullString
    Next
    v.size = n
End Sub

' C: v.clear()
Public Sub StringVector_clear(v As StringVector)
    v.size = 0
    Erase v.items
End Sub
