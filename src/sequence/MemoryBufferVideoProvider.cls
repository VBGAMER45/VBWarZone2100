VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryBufferVideoProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/sequence/sequence.cpp: class MemoryBufferVideoProvider : public VideoProvider, a video from a memory
' buffer (seqdisp.cpp's on-demand downloads). Private: a VideoProvider holds it (makeVideoProvider_data). The
' std::shared_ptr<const std::vector<char>> is a copy of the bytes (hasBuffer: != nullptr).
' buffer_data's memcpy copies `size` (256K) bytes from currPos whatever is left (it reads past the vector's end
' on the last chunk): here only the bytes there are, which are all that ogg_sync_wrote counts.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private memoryBuffer() As Byte          ' std::shared_ptr<const std::vector<char>> memoryBuffer
Private hasBuffer As Boolean            ' memoryBuffer != nullptr
Private bufSize As Long                 ' memoryBuffer->size()
Private currPos As Long                 ' size_t currPos = 0

' MemoryBufferVideoProvider(std::shared_ptr<const std::vector<char>> memoryBuffer, const WzString& filename):
' data(first .. first + n - 1); n < 0 for nullptr
Friend Sub construct(data() As Byte, ByVal first As Long, ByVal n As Long)
    If n < 0 Then
        hasBuffer = False
        bufSize = 0
        Exit Sub
    End If
    hasBuffer = True
    bufSize = n
    If n > 0 Then
        ReDim memoryBuffer(0 To n - 1)
        CopyMemory memoryBuffer(0), data(first), n
    End If
End Sub

Friend Function seek_begin() As Boolean
    currPos = 0
    seek_begin = True
End Function

Friend Function buffer_data(oy As ogg_sync_state) As Long
    Const size As Long = 262144
    Dim buffer As Long, bytesToRead As Long, bytes As Long
    If Not hasBuffer Then buffer_data = 0: Exit Function
    If bufSize = 0 Then buffer_data = 0: Exit Function      ' memoryBuffer->empty()
    If currPos >= bufSize Then buffer_data = 0: Exit Function

    ' read in 256K chunks
    buffer = ogg_sync_buffer(oy, size)  ' (VB) the offset into oy.data

    bytesToRead = bufSize - currPos     ' std::min<size_t>(size, memoryBuffer->size() - currPos)
    If bytesToRead > size Then bytesToRead = size
    CopyMemory oy.data(buffer), memoryBuffer(currPos), bytesToRead    ' (memcpy(buffer, ..., size): see above)
    currPos = currPos + bytesToRead

    bytes = bytesToRead

    ogg_sync_wrote oy, bytes
    buffer_data = (bytes)
End Function

Friend Function end_of_data() As Boolean
    If Not hasBuffer Then end_of_data = True: Exit Function
    end_of_data = currPos >= bufSize
End Function
