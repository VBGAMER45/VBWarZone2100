VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectionWriteQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/pending_writes_manager.h: using ConnectionWriteQueue = std::vector<uint8_t>, as an entry of
' PendingWritesManager's ConnectionThreadWriteMap (std::unordered_map<IClientConnection*, ConnectionWriteQueue>): the
' bytes waiting to be sent, and the connection (the map's key). The array grows by doubling (a vector's
' amortised append); erasing the front moves the rest down, as vector::erase does.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private conn_ As IClientConnection
Private data_() As Byte
Private size_ As Long
Private cap_ As Long

Friend Sub construct(ByVal conn As IClientConnection)
    Set conn_ = conn
End Sub

' (VB) the map entry's key
Friend Function conn() As IClientConnection
    Set conn = conn_
End Function

' writeQueue.insert(writeQueue.end(), buf, buf + size)
Friend Sub append(buf() As Byte, ByVal ofs As Long, ByVal n As Long)
    Dim newCap As Long
    If n <= 0 Then Exit Sub
    If size_ + n > cap_ Then
        newCap = cap_ * 2
        If newCap < size_ + n Then newCap = size_ + n
        If newCap < 256 Then newCap = 256
        If cap_ = 0 Then
            ReDim data_(0 To newCap - 1)
        Else
            ReDim Preserve data_(0 To newCap - 1)
        End If
        cap_ = newCap
    End If
    CopyMemory data_(size_), buf(ofs), n
    size_ = size_ + n
End Sub

' writeQueue.erase(writeQueue.begin(), writeQueue.begin() + n)
Friend Sub eraseFront(ByVal n As Long)
    If n >= size_ Then
        size_ = 0
    ElseIf n > 0 Then
        CopyMemory data_(0), data_(n), size_ - n
        size_ = size_ - n
    End If
End Sub

' writeQueue.empty_()
Friend Function empty_() As Boolean
    empty_ = (size_ = 0)
End Function

' writeQueue.size()
Friend Function size() As Long
    size = size_
End Function

' (VB) conn->sendImpl(writeQueue): the bytes go to the connection from here (VB cannot hand out the array)
Friend Function sendTo(ByVal conn As IClientConnection, ec As Long) As Long
    sendTo = conn.sendImpl(data_, size_, ec)
End Function
