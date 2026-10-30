VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMapBinaryPhysFSStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/map.cpp: class WzMapBinaryPhysFSStream, the stream of WzMapPhysFSIO
' (a PhysFS file). construct is the constructor; the destructor closes it
' (Class_Terminate). WZ_PHYSFS_SETBUFFER (PhysFS's read buffer) is not
' ported: it changes no result.
Option Explicit

Implements BinaryIOStream

Private pFile As Long                   ' PHYSFS_File *

Friend Sub construct(ByVal filename As String, ByVal mode As OpenMode)
    Select Case mode
    Case OpenMode_READ
        pFile = PHYSFS_openRead(filename)
    Case OpenMode_WRITE
        pFile = PHYSFS_openWrite(filename)
    End Select
End Sub

Private Sub Class_Terminate()
    BinaryIOStream_closeStream
End Sub

Friend Function openedFile() As Boolean
    openedFile = pFile <> 0
End Function

Private Function BinaryIOStream_readBytes(buffer() As Byte, ByVal offset As Long, ByVal len_ As Long, bytesRead As Long) As Boolean
    Dim result As Long
    If pFile = 0 Then BinaryIOStream_readBytes = False: Exit Function
    result = PHYSFS_readBytes(pFile, buffer, offset, len_)
    If result < 0 Then
        ' failed
        BinaryIOStream_readBytes = False
        Exit Function
    End If
    bytesRead = result
    BinaryIOStream_readBytes = True
End Function

Private Function BinaryIOStream_writeBytes(buffer() As Byte, ByVal offset As Long, ByVal len_ As Long, bytesWritten As Long) As Boolean
    Dim result As Long
    If pFile = 0 Then BinaryIOStream_writeBytes = False: Exit Function
    result = PHYSFS_writeBytes(pFile, buffer, offset, len_)
    If result < 0 Then
        ' failed
        BinaryIOStream_writeBytes = False
        Exit Function
    End If
    bytesWritten = result
    BinaryIOStream_writeBytes = True
End Function

Private Function BinaryIOStream_closeStream() As Boolean
    If pFile = 0 Then
        BinaryIOStream_closeStream = False
        Exit Function
    End If

    PHYSFS_close pFile
    pFile = 0
    BinaryIOStream_closeStream = True
End Function

Private Function BinaryIOStream_endOfStream() As Boolean
    If pFile = 0 Then BinaryIOStream_endOfStream = False: Exit Function
    BinaryIOStream_endOfStream = PHYSFS_eof(pFile) <> 0
End Function
