Attribute VB_Name = "stdio_win"
' C's stdio for the port, on Win32 (no C++ counterpart; compiled into every
' project, like a static C runtime). A FILE* is an index into files()
' (0 = NULL); STDERR_FILE is stderr, STDOUT_FILE stdout. They write to the
' process's standard handles, looked up at each write, so a freopen in one
' project redirects them in all (Warzone has one C library). Text is written as UTF-8, and a file
' opened without "b" is in text mode, where each LF becomes CR LF, as in the
' Microsoft C library Warzone's Windows build uses. Paths are full Unicode
' (CreateFileW), like Warzone's _wfopen.
Option Explicit

Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function SetStdHandle Lib "kernel32" (ByVal nStdHandle As Long, ByVal hHandle As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = 1
Private Const FILE_SHARE_WRITE As Long = 2
Private Const CREATE_ALWAYS As Long = 2
Private Const OPEN_EXISTING As Long = 3
Private Const OPEN_ALWAYS As Long = 4
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_END As Long = 2
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const STD_OUTPUT_HANDLE As Long = -11
Private Const STD_ERROR_HANDLE As Long = -12
Private Const CP_UTF8 As Long = 65001

Public Const STDERR_FILE As Long = 1
Public Const STDOUT_FILE As Long = 2

Private Type FILE_t
    handle As Long
    textMode As Boolean
    inUse As Boolean
End Type

Private files() As FILE_t
Private filesReady As Boolean

Private Sub InitFiles()
    ReDim files(1 To 16)
    files(STDERR_FILE).textMode = True
    files(STDERR_FILE).inUse = True
    files(STDOUT_FILE).textMode = True
    files(STDOUT_FILE).inUse = True
    filesReady = True
End Sub

' The OS handle of f; stdout and stderr are the process's (0 or -1 without a console: writes go nowhere).
Private Function FileHandle(ByVal f As Long) As Long
    If f = STDERR_FILE Then
        FileHandle = GetStdHandle(STD_ERROR_HANDLE)
    ElseIf f = STDOUT_FILE Then
        FileHandle = GetStdHandle(STD_OUTPUT_HANDLE)
    Else
        FileHandle = files(f).handle
    End If
End Function

' C: fopen (path, mode); mode "r", "w" or "a", with "b" for binary and "+" for update. 0 = NULL.
Public Function fopen(ByVal path As String, ByVal mode As String) As Long
    Dim i As Long, h As Long, access As Long, disp As Long
    If Not filesReady Then InitFiles
    Select Case Left$(mode, 1)
    Case "r"
        access = GENERIC_READ
        disp = OPEN_EXISTING
    Case "w"
        access = GENERIC_WRITE
        disp = CREATE_ALWAYS
    Case "a"
        access = GENERIC_WRITE
        disp = OPEN_ALWAYS
    Case Else
        Exit Function
    End Select
    If InStr(mode, "+") > 0 Then access = GENERIC_READ Or GENERIC_WRITE
    h = CreateFileW(StrPtr(path), access, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, disp, FILE_ATTRIBUTE_NORMAL, 0)
    If h = INVALID_HANDLE_VALUE Then Exit Function
    If Left$(mode, 1) = "a" Then SetFilePointer h, 0, 0, FILE_END
    For i = STDOUT_FILE + 1 To UBound(files)
        If Not files(i).inUse Then Exit For
    Next
    If i > UBound(files) Then ReDim Preserve files(1 To i + 15)
    files(i).handle = h
    files(i).textMode = (InStr(mode, "b") = 0)
    files(i).inUse = True
    fopen = i
End Function

' C: fputs (s, f): writes s as UTF-8. Returns 0, or -1 (EOF) on failure.
Public Function fputs(ByVal s As String, ByVal f As Long) As Long
    Dim n As Long, b() As Byte, written As Long
    If Not filesReady Then InitFiles
    If f < 1 Or f > UBound(files) Then
        fputs = -1
        Exit Function
    End If
    If Not files(f).inUse Then
        fputs = -1
        Exit Function
    End If
    If Len(s) = 0 Then Exit Function
    If files(f).textMode Then s = Replace(s, vbLf, vbCrLf)
    n = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
    If n <= 0 Then
        fputs = -1
        Exit Function
    End If
    ReDim b(0 To n - 1)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(s), Len(s), VarPtr(b(0)), n, 0, 0
    If WriteFile(FileHandle(f), b(0), n, written, 0) = 0 Then fputs = -1
End Function

' C: fprintf (f, fmt, ...)
Public Function fprintf(ByVal f As Long, ByVal fmt As String, ParamArray args() As Variant) As Long
    Dim a As Variant, s As String
    a = args
    s = vsprintf(fmt, a)
    If fputs(s, f) = 0 Then fprintf = Len(s) Else fprintf = -1
End Function

' C: fflush (f)
Public Function fflush(ByVal f As Long) As Long
    If Not filesReady Then InitFiles
    If f < 1 Or f > UBound(files) Then Exit Function
    If files(f).inUse Then FlushFileBuffers FileHandle(f)
End Function

' C: freopen (path, mode, f), for stdout and stderr (mode "w"): the file
' becomes the process's standard handle. "CONOUT$" is the console.
Public Function freopen(ByVal path As String, ByVal mode As String, ByVal f As Long) As Long
    Dim h As Long, disp As Long
    If Not filesReady Then InitFiles
    If (f <> STDERR_FILE And f <> STDOUT_FILE) Or Left$(mode, 1) <> "w" Then Exit Function
    If UCase$(path) = "CONOUT$" Then disp = OPEN_EXISTING Else disp = CREATE_ALWAYS
    h = CreateFileW(StrPtr(path), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, disp, FILE_ATTRIBUTE_NORMAL, 0)
    If h = INVALID_HANDLE_VALUE Then Exit Function
    If f = STDERR_FILE Then SetStdHandle STD_ERROR_HANDLE, h Else SetStdHandle STD_OUTPUT_HANDLE, h
    files(f).textMode = (InStr(mode, "b") = 0)
    freopen = f
End Function

' C: fclose (f)
Public Function fclose(ByVal f As Long) As Long
    If Not filesReady Then InitFiles
    If f <= STDOUT_FILE Or f > UBound(files) Then
        fclose = -1
        Exit Function
    End If
    If Not files(f).inUse Then
        fclose = -1
        Exit Function
    End If
    CloseHandle files(f).handle
    files(f).inUse = False
    files(f).handle = 0
End Function
