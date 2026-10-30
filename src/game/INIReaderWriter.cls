VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INIReaderWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/3rdparty/INIReaderWriter.cpp: class INIReaderWriter. The sections
' (a std::map with the case-insensitive comparator) are kept sorted by
' ascii_strcasecmp, so dump writes them in C++'s order. A new object is the
' default constructor; InitFromData is the (buffer, size) constructor. The
' (filename) constructor is not ported (Warzone reads through PhysFS).
Option Explicit

Private sectionNames() As String
Private sections() As IniSection
Private sectionCount As Long
Private error_ As Long
Private loaded_ As Boolean

' C: INIReaderWriter(const char *buffer, size_t buffer_size): buffer(0 .. buffer_size - 1)
Friend Sub InitFromData(buffer() As Byte, ByVal buffer_size As Long)
    Dim n As Long
    ' string content(buffer, buffer_size); ini_parse_string reads it up to its first NUL
    n = 0
    Do While n < buffer_size
        If buffer(n) = 0 Then Exit Do
        n = n + 1
    Loop
    error_ = ini_parse_string(buffer, n, Me)
    loaded_ = True
End Sub

Public Function ParseError() As Long
    ParseError = error_
End Function

Public Function LoadedFromData() As Boolean
    LoadedFromData = loaded_
End Function

Private Function FindSection(ByVal section As String) As Long
    Dim lo As Long, hi As Long, m As Long, c As Long
    lo = 0
    hi = sectionCount - 1
    Do While lo <= hi
        m = (lo + hi) \ 2
        c = ascii_strcasecmp(sectionNames(m), section)
        If c = 0 Then
            FindSection = m
            Exit Function
        End If
        If c < 0 Then lo = m + 1 Else hi = m - 1
    Loop
    FindSection = -lo - 1
End Function

' C: IniSection& GetSection (std::map's operator[]: made if missing)
Public Function GetSection(ByVal section As String) As IniSection
    Dim i As Long, k As Long
    i = FindSection(section)
    If i < 0 Then
        i = -i - 1
        ReDim Preserve sectionNames(0 To sectionCount)
        ReDim Preserve sections(0 To sectionCount)
        For k = sectionCount To i + 1 Step -1
            sectionNames(k) = sectionNames(k - 1)
            Set sections(k) = sections(k - 1)
        Next
        sectionNames(i) = section
        Set sections(i) = New IniSection
        sectionCount = sectionCount + 1
    End If
    Set GetSection = sections(i)
End Function

Public Function HasSection(ByVal section As String) As Boolean
    HasSection = (FindSection(section) >= 0)
End Function

Public Function HasValue(ByVal section As String, ByVal name_ As String) As Boolean
    Dim i As Long
    i = FindSection(section)
    If i < 0 Then Exit Function
    HasValue = sections(i).HasValue(name_)
End Function

' C: static ValueHandler (user, section, name, value), called by ini_parse_string
Friend Function ValueHandler(ByVal section As String, ByVal name_ As String, ByVal value As String) As Long
    GetSection(section).SetString name_, value
    ValueHandler = 1
End Function

' Dump the current ini state to a string (ex. for output to a file)
Public Function dump() As String
    Dim endl As String, result As String, i As Long
    endl = vbCrLf                       ' _WIN32
    For i = 0 To sectionCount - 1
        result = result & "[" & sectionNames(i) & "]" & endl
        sections(i).dumpTo result, endl
    Next
    dump = result
End Function
