VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/3rdparty/INIReaderWriter.cpp: INIReaderWriter::IniSection. Its
' std::map with INIDetail::CaseInsensitiveComparator is two arrays kept
' sorted by ascii_strcasecmp, so dumpTo writes the keys in C++'s order.
' C++'s Get is Get_ (Get is a VB statement).
Option Explicit

Private names_() As String
Private values_() As String
Private count_ As Long

' The index of name, or -(insertion point) - 1 when it is not there.
Private Function Find(ByVal name_ As String) As Long
    Dim lo As Long, hi As Long, m As Long, c As Long
    lo = 0
    hi = count_ - 1
    Do While lo <= hi
        m = (lo + hi) \ 2
        c = ascii_strcasecmp(names_(m), name_)
        If c = 0 Then
            Find = m
            Exit Function
        End If
        If c < 0 Then lo = m + 1 Else hi = m - 1
    Loop
    Find = -lo - 1
End Function

Public Function HasValue(ByVal name_ As String) As Boolean
    HasValue = (Find(name_) >= 0)
End Function

Public Function has(ByVal name_ As String) As Boolean
    has = HasValue(name_)
End Function

Public Function Get_(ByVal name_ As String, ByVal default_value As String) As String
    Dim i As Long
    i = Find(name_)
    If i < 0 Then
        Get_ = default_value
        Exit Function
    End If
    Get_ = values_(i)
End Function

Public Function GetString(ByVal name_ As String, ByVal default_value As String) As String
    Dim str_ As String
    str_ = Get_(name_, "")
    If Len(str_) = 0 Then GetString = default_value Else GetString = str_
End Function

' This parses "1234" (decimal) and also "0x4D2" (hex)
Public Function GetInteger(ByVal name_ As String, ByVal default_value As Long) As Long
    Dim valstr As String, n As Long, endIdx As Long
    valstr = Get_(name_, "")
    n = strtol(valstr, endIdx, 0)
    If endIdx > 0 Then GetInteger = n Else GetInteger = default_value
End Function

Public Function GetInteger64(ByVal name_ As String, ByVal default_value As Double) As Double
    Dim valstr As String, n As Double, endIdx As Long
    valstr = Get_(name_, "")
    n = strtoll(valstr, endIdx, 0)
    If endIdx > 0 Then GetInteger64 = n Else GetInteger64 = default_value
End Function

Public Function GetUnsigned(ByVal name_ As String, ByVal default_value As Long) As Long
    Dim valstr As String, n As Long, endIdx As Long
    valstr = Get_(name_, "")
    n = strtoul(valstr, endIdx, 0)
    If endIdx > 0 Then GetUnsigned = n Else GetUnsigned = default_value
End Function

Public Function GetReal(ByVal name_ As String, ByVal default_value As Double) As Double
    Dim valstr As String, n As Double, endIdx As Long
    valstr = Get_(name_, "")
    n = strtod(valstr, endIdx)
    If endIdx > 0 Then GetReal = n Else GetReal = default_value
End Function

Public Function GetBoolean(ByVal name_ As String, ByVal default_value As Boolean) As Boolean
    Dim valstr As String
    ' Convert to lower case to make string comparisons case-insensitive
    valstr = LCase$(Get_(name_, ""))
    If valstr = "true" Or valstr = "yes" Or valstr = "on" Or valstr = "1" Then
        GetBoolean = True
    ElseIf valstr = "false" Or valstr = "no" Or valstr = "off" Or valstr = "0" Then
        GetBoolean = False
    Else
        GetBoolean = default_value
    End If
End Function

' Set a string value in a section
Public Sub SetString(ByVal name_ As String, ByVal value As String)
    Dim i As Long, k As Long
    i = Find(name_)
    If i >= 0 Then
        values_(i) = value
        Exit Sub
    End If
    i = -i - 1
    ReDim Preserve names_(0 To count_)
    ReDim Preserve values_(0 To count_)
    For k = count_ To i + 1 Step -1
        names_(k) = names_(k - 1)
        values_(k) = values_(k - 1)
    Next
    names_(i) = name_
    values_(i) = value
    count_ = count_ + 1
End Sub

Public Sub SetBool(ByVal name_ As String, ByVal value As Boolean)
    If value Then SetString name_, "true" Else SetString name_, "false"
End Sub

Public Sub dumpTo(outputStr As String, ByVal endl As String)
    Dim i As Long
    For i = 0 To count_ - 1
        outputStr = outputStr & names_(i) & "=" & values_(i) & endl
    Next
End Sub
