Attribute VB_Name = "prof_vb"
' (VB-only) section timers for profiling: profBegin k / profEnd k around a
' section (k = 0 .. 15), profWrite <file> writes each one's count and
' seconds. Calls are added for a measurement and removed after; none stay
' in the sources.
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Public profTotal(0 To 15) As Currency
Public profCount(0 To 15) As Long
Private profStart(0 To 15) As Currency

Public Sub profBegin(ByVal k As Long)
    QueryPerformanceCounter profStart(k)
End Sub

Public Sub profEnd(ByVal k As Long)
    Dim t As Currency
    QueryPerformanceCounter t
    profTotal(k) = profTotal(k) + (t - profStart(k))
    profCount(k) = profCount(k) + 1
End Sub

Public Sub profWrite(ByVal path As String)
    Dim f As Integer, k As Long, freq As Currency
    QueryPerformanceFrequency freq
    f = FreeFile
    Open path For Output As #f
    For k = 0 To 15
        If profCount(k) > 0 Then Print #f, k, profCount(k), Format$(profTotal(k) / freq, "0.000")
    Next
    Close #f
End Sub
