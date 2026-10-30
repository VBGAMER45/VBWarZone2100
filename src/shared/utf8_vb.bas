Attribute VB_Name = "utf8_vb"
' UTF-8 for the port (no C++ counterpart; compiled into every project).
' Warzone's text is UTF-8 in std::string and WzString; a VB String is
' UTF-16, so text is converted where it meets bytes: files, the network.
Option Explicit

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Private Const CP_UTF8 As Long = 65001

' s as UTF-8 in b(0 .. n - 1); returns n (b is left alone when s is empty).
Public Function utf8_encode(ByVal s As String, b() As Byte) As Long
    Dim n As Long
    If Len(s) = 0 Then Exit Function
    n = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
    If n <= 0 Then Exit Function
    ReDim b(0 To n - 1)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(s), Len(s), VarPtr(b(0)), n, 0, 0
    utf8_encode = n
End Function

' The String of the UTF-8 bytes b(ofs .. ofs + n - 1).
Public Function utf8_decode(b() As Byte, ByVal ofs As Long, ByVal n As Long) As String
    Dim w As Long, s As String
    If n <= 0 Then Exit Function
    w = MultiByteToWideChar(CP_UTF8, 0, VarPtr(b(ofs)), n, 0, 0)
    If w <= 0 Then Exit Function
    s = String$(w, 0)
    MultiByteToWideChar CP_UTF8, 0, VarPtr(b(ofs)), n, StrPtr(s), w
    utf8_decode = s
End Function
