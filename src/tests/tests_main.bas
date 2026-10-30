Attribute VB_Name = "tests_main"
' WzTests.exe (VB-only, not part of the game): runs the commands of a
' command file through the port's code and writes their output, which
' tools/run_tests.py compares with the C++ reference (tools/jsonref.cpp).
'
'   WzTests <command file> <output folder>
'
' Each line: mode <TAB> input file <TAB> output file (relative to the output
' folder), with jsonref's modes (dump, dumpo, compact, ascii, error, doubles)
' WzConfig's (wzread, wzcopy: test_wzconfig) and wzmaplib's (map, export:
' test_map, against tools/mapref.cpp).
Option Explicit

Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Private Function GetArgs(argv() As String) As Long
    Dim p As Long, i As Long, sp As Long, n As Long, argc As Long
    p = CommandLineToArgvW(GetCommandLineW(), argc)
    ReDim argv(0 To argc)
    For i = 0 To argc - 1
        CopyMemory sp, ByVal p + i * 4, 4
        n = lstrlenW(sp)
        argv(i) = String$(n, 0)
        If n > 0 Then CopyMemory ByVal StrPtr(argv(i)), ByVal sp, n * 2
    Next
    LocalFree p
    GetArgs = argc
End Function

Public Function ReadBytes(ByVal path As String, b() As Byte) As Long
    Dim f As Integer, n As Long
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    Else
        ReDim b(0 To 0)
    End If
    Close #f
    ReadBytes = n
End Function

Public Sub WriteText(ByVal path As String, ByVal s As String)
    Dim f As Integer, b() As Byte, n As Long
    n = utf8_encode(s, b)
    f = FreeFile
    If Len(Dir$(path)) > 0 Then Kill path
    Open path For Binary Access Write As #f
    If n > 0 Then Put #f, , b
    Close #f
End Sub

Sub Main()
    Dim argv() As String, argc As Long, lines() As String, i As Long, parts() As String
    Dim b() As Byte, n As Long, root As String, done As Long
    argc = GetArgs(argv)
    If argc < 3 Then ExitProcess 2
    n = ReadBytes(argv(1), b)
    lines = Split(Replace(utf8_decode(b, 0, n), vbCr, ""), vbLf)
    root = argv(2)
    For i = 0 To UBound(lines)
        parts = Split(lines(i), vbTab)
        If UBound(parts) = 2 Then
            If parts(0) = "map" Then
                WriteText root & "\" & Replace(parts(2), "/", "\"), test_map.RunMap(parts(0), parts(1), root, parts(2))
            ElseIf parts(0) = "export" Then
                WriteText root & "\" & Replace(parts(2), "/", "\") & ".log", test_map.RunMap(parts(0), parts(1), root, parts(2))
            ElseIf Left$(parts(0), 2) = "wz" Then
                WriteText root & "\" & Replace(parts(2), "/", "\"), test_wzconfig.RunWzConfig(parts(0), parts(1), root)
            Else
                WriteText root & "\" & Replace(parts(2), "/", "\"), test_json.RunJson(parts(0), parts(1))
            End If
            done = done + 1
        End If
    Next
    ExitProcess 0
End Sub
