VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WZAudioDataMemoryBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/sound/codecs.h: class WZAudioDataMemoryBuffer : public WZAudioDataResourceInterface, a whole file in
' memory (oggopus.cpp's wz_opus_membuf_* callbacks). (VB) A class has no public array: std::vector<char> buffer
' is set with vbSetBuffer, read with bufferSize (buffer.size()) and vbCopyOut (a memcpy from it).
Option Explicit

Implements WZAudioDataResourceInterface

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private buffer() As Byte                        ' std::vector<char> buffer
Private bufferCount As Long                     ' (VB) buffer.size()
Public currPos As Long                          ' size_t = 0
Public error_ As Boolean                        ' bool error = false

' (VB) the vector's contents: data(0 .. size - 1)
Friend Sub vbSetBuffer(data() As Byte, ByVal size As Long)
    bufferCount = size
    If size > 0 Then
        ReDim buffer(0 To size - 1)
        CopyMemory buffer(0), data(0), size
    Else
        Erase buffer
    End If
End Sub

Friend Function bufferSize() As Long
    bufferSize = bufferCount
End Function

' (VB) memcpy(dst + dstOff, &buffer[srcPos], n)
Friend Sub vbCopyOut(dst() As Byte, ByVal dstOff As Long, ByVal srcPos As Long, ByVal n As Long)
    If n > 0 Then CopyMemory dst(dstOff), buffer(srcPos), n
End Sub
