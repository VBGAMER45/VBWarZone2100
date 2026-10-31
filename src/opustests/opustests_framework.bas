Attribute VB_Name = "opustests_framework"
' (VB-only) OpusTests.exe's stand-ins for what WzSound takes from WzFramework (the test exe references no DLL):
' the log (debug_, ASSERT_FAILURE, code_part's LOG_* constants: the lines go to <output folder>\log.txt), and
' the PhysFS calls oggopus.cpp makes, on the real files (a path is a file path; PHYSFS_file * is a Win32
' handle), with PhysFS's results: PHYSFS_readBytes's count, PHYSFS_seek's non-zero on success,
' PHYSFS_fileLength and PHYSFS_tell (-1 on an error), loadFileToBufferVector (frame.cpp) the whole file.
Option Explicit

Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' lib/framework/debug.h: enum code_part (the parts the tested modules log to)
Public Const LOG_ALL As Long = 0
Public Const LOG_MAIN As Long = 1
Public Const LOG_SOUND As Long = 2
Public Const LOG_VIDEO As Long = 3
Public Const LOG_WZ As Long = 4
Public Const LOG_ERROR As Long = 10
Public Const LOG_WARNING As Long = 11
Public Const LOG_NEVER As Long = 12
Public Const LOG_INFO As Long = 13
Public Const LOG_FATAL As Long = 14

Public logPath As String                ' (set by Main)
Public logCount As Long                 ' log lines written
Public assertCount As Long              ' ASSERT_FAILUREs

Private Sub logLine(ByVal s As String)
    Dim f As Integer
    logCount = logCount + 1
    If Len(logPath) = 0 Then Exit Sub
    f = FreeFile
    Open logPath For Append As #f
    Print #f, s
    Close #f
End Sub

Public Sub debug_(ByVal line_ As Long, ByVal part As Long, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    logLine "debug " & part & " " & function_ & ":" & line_ & ": " & vsprintf(str_, a)
End Sub

Public Function debugPartEnabled(ByVal codePart As Long) As Boolean
    debugPartEnabled = True
End Function

Public Sub ASSERT_FAILURE(ByVal expr_string As String, ByVal location_description As String, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    assertCount = assertCount + 1
    logLine "ASSERT " & location_description & " " & function_ & ": " & expr_string & ": " & vsprintf(str_, a)
End Sub

Public Function PHYSFS_openRead(ByVal fname As String) As Long
    Dim h As Long
    h = CreateFileW(StrPtr(fname), &H80000000, 3, 0, 3, &H80, 0)
    If h = -1 Then h = 0
    PHYSFS_openRead = h
End Function

Public Function WZ_PHYSFS_readBytes(ByVal handle As Long, buf() As Byte, ByVal first As Long, ByVal size As Long) As Double
    Dim got As Long
    If size <= 0 Then Exit Function
    If ReadFile(handle, buf(first), size, got, 0) = 0 Then WZ_PHYSFS_readBytes = -1 Else WZ_PHYSFS_readBytes = got
End Function

Public Function PHYSFS_seek(ByVal f As Long, ByVal pos As Double) As Long
    Dim hi As Long, lo As Long, l As Double
    If pos < 0 Then Exit Function
    hi = Int(pos / 4294967296#)
    l = pos - hi * 4294967296#
    If l >= 2147483648# Then lo = l - 4294967296# Else lo = l
    If SetFilePointer(f, lo, hi, 0) = -1 And Err.LastDllError <> 0 Then Exit Function
    PHYSFS_seek = 1
End Function

Public Function PHYSFS_tell(ByVal f As Long) As Double
    Dim hi As Long, lo As Long
    lo = SetFilePointer(f, 0, hi, 1)
    PHYSFS_tell = hi * 4294967296# + IIf(lo < 0, lo + 4294967296#, CDbl(lo))
End Function

Public Function PHYSFS_fileLength(ByVal f As Long) As Double
    Dim hi As Long, lo As Long
    lo = GetFileSize(f, hi)
    PHYSFS_fileLength = hi * 4294967296# + IIf(lo < 0, lo + 4294967296#, CDbl(lo))
End Function

Public Function PHYSFS_close(ByVal f As Long) As Long
    If CloseHandle(f) <> 0 Then PHYSFS_close = 1
End Function

Public Function PHYSFS_getRealDir(ByVal fname As String) As String
    PHYSFS_getRealDir = "(test)"
End Function

Public Function WZ_PHYSFS_getLastError() As String
    WZ_PHYSFS_getLastError = "not found"
End Function

Public Function loadFileToBufferVector(ByVal pFileName As String, outputBuffer() As Byte, outputSize As Long, ByVal hard_fail As Boolean, ByVal appendNullCharacter As Boolean) As Boolean
    Dim h As Long, n As Double
    h = PHYSFS_openRead(pFileName)
    If h = 0 Then Exit Function
    n = PHYSFS_fileLength(h)
    If n > 0 Then
        ReDim outputBuffer(0 To CLng(n) - 1)
        WZ_PHYSFS_readBytes h, outputBuffer, 0, CLng(n)
    Else
        ReDim outputBuffer(0 To 0)
    End If
    outputSize = n
    PHYSFS_close h
    loadFileToBufferVector = True
End Function
