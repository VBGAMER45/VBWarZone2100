VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMapBinaryStdIOStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/wzmaplib/src/map_io.cpp: class WzMapBinaryStdIOStream, the stream of
' StdIOProvider, which is a C FILE* opened with _wfopen. Here it is the
' Windows file API with the C library's behaviour kept:
' * "rb": the file is read whole when it is opened, and fread copies from
'   that copy. feof is set by the read that comes up short; fsetpos clears
'   it (endOfStream).
' * "wb": the file is created or truncated and written through a buffer,
'   as stdio does; the buffer is flushed when full and at fclose.
' * A read error or a write error sets ferror, which stays set.
' Made by StdIOProvider's openBinaryStream (construct is the constructor);
' the destructor closes it (Class_Terminate).
Option Explicit

Implements BinaryIOStream

Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = 1
Private Const FILE_SHARE_WRITE As Long = 2
Private Const CREATE_ALWAYS As Long = 2
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const WRITE_BUFFER_SIZE As Long = 4096

Private pFileOpen As Boolean            ' pFile != nullptr
Private hFile As Long                   ' "wb": the open file
Private writing As Boolean
Private mData() As Byte                 ' "rb": the file's bytes
Private mSize As Long
Private mPos As Long                    ' the file position
Private eofFlag As Boolean              ' feof
Private errFlag As Boolean              ' ferror
Private wbuf() As Byte                  ' "wb": bytes not yet written
Private wlen As Long

' WzMapBinaryStdIOStream(const char* pFilename, const char mode): mode "r" or "w"
Friend Sub construct(ByVal pFilename As String, ByVal mode As String)
    Dim h As Long, n As Long, hi As Long, got As Long
    If mode = "w" Then
        h = CreateFileW(StrPtr(pFilename), GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
        If h = INVALID_HANDLE_VALUE Then Exit Sub
        hFile = h
        writing = True
        ReDim wbuf(0 To WRITE_BUFFER_SIZE - 1)
        pFileOpen = True
        Exit Sub
    End If
    h = CreateFileW(StrPtr(pFilename), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If h = INVALID_HANDLE_VALUE Then Exit Sub
    n = GetFileSize(h, hi)
    If n > 0 Then
        ReDim mData(0 To n - 1)
        If ReadFile(h, mData(0), n, got, 0) = 0 Then
            errFlag = True
            got = 0
        End If
        mSize = got
    End If
    CloseHandle h
    pFileOpen = True
End Sub

Public Function openedFile() As Boolean
    openedFile = pFileOpen
End Function

Private Function Flush() As Boolean
    Dim written As Long
    If wlen > 0 Then
        If WriteFile(hFile, wbuf(0), wlen, written, 0) = 0 Or written <> wlen Then errFlag = True
        wlen = 0
    End If
    Flush = Not errFlag
End Function

Private Function BinaryIOStream_readBytes(buffer() As Byte, ByVal offset As Long, ByVal len_ As Long, bytesRead As Long) As Boolean
    Dim n As Long
    bytesRead = 0
    If eofFlag Then
        BinaryIOStream_readBytes = True
        Exit Function
    End If
    If writing Then
        errFlag = True
    Else
        n = mSize - mPos
        If n > len_ Then n = len_
        If n < len_ Then eofFlag = True
        If n > 0 Then CopyMemory buffer(offset), mData(mPos), n
        mPos = mPos + n
        bytesRead = n
    End If
    If errFlag Then
        Exit Function
    End If
    BinaryIOStream_readBytes = True
End Function

Private Function BinaryIOStream_writeBytes(buffer() As Byte, ByVal offset As Long, ByVal len_ As Long, bytesWritten As Long) As Boolean
    bytesWritten = 0
    If Not writing Then
        errFlag = True
        Exit Function
    End If
    If len_ > 0 Then
        If wlen + len_ > WRITE_BUFFER_SIZE Then Flush
        If len_ > WRITE_BUFFER_SIZE Then
            ReDim Preserve wbuf(0 To len_ - 1)
        End If
        CopyMemory wbuf(wlen), buffer(offset), len_
        wlen = wlen + len_
        If len_ > WRITE_BUFFER_SIZE Then
            Flush
            ReDim wbuf(0 To WRITE_BUFFER_SIZE - 1)
        End If
        mPos = mPos + len_
        bytesWritten = len_
    End If
    If errFlag Then
        Exit Function
    End If
    BinaryIOStream_writeBytes = True
End Function

Private Function BinaryIOStream_closeStream() As Boolean
    If Not pFileOpen Then
        BinaryIOStream_closeStream = False
        Exit Function
    End If
    If writing Then
        Flush
        CloseHandle hFile
        hFile = 0
    End If
    Erase mData
    pFileOpen = False
    BinaryIOStream_closeStream = True
End Function

Private Function BinaryIOStream_endOfStream() As Boolean
    Dim endOfStream As Boolean, position As Long, tmpByte As Long, s As BinaryIOStream
    endOfStream = eofFlag
    If Not endOfStream Then
        ' feof may only be set once a read fails
        position = mPos                     ' fgetpos
        ' attempt to read a byte
        Set s = Me
        If Not BinaryIOStream_readULE8(s, tmpByte) Then
            ' read failed - check feof again
            endOfStream = eofFlag
        End If
        ' restore original position
        mPos = position                     ' fsetpos (which clears the end-of-file indicator)
        eofFlag = False
    End If
    BinaryIOStream_endOfStream = endOfStream
End Function

' ~WzMapBinaryStdIOStream()
Private Sub Class_Terminate()
    BinaryIOStream_closeStream
End Sub
