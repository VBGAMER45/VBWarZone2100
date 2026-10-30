VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MessageReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/netplay/netqueue.h: class MessageReader, used for deserialising, using
' the same interface as MessageWriter. Only NETbeginDecode makes one
' (netqueue_c's MessageReader_new); the exe holds it and passes it to
' NetApi's NET* procedures. It keeps the message (C++: a pointer to its
' data; messages are immutable, so the object is shared) and the read
' index; its members are for WzNet only (Friend).
Option Explicit

Private msgData As NetMessage           ' const NetMsgDataVector* msgData
Private index As Long                   ' mutable size_t index = NetMessage::HEADER_LENGTH
Private msgType_ As Long                ' uint8_t msgType

Friend Sub construct(ByVal m As NetMessage)
    Set msgData = m
    index = NetMessage_HEADER_LENGTH
    msgType_ = m.type_()
End Sub

Friend Property Get msgType() As Long
    msgType = msgType_
End Property

' byte(uint8_t &v)
Friend Sub byte_(v As Long)
    If index >= msgData.rawDataSize() Then
        v = 0
    Else
        v = msgData.dataByte(index)
    End If
    index = index + 1
End Sub

' bytes(uint8_t *pOut, size_t numBytes): pOut(ofs .. ofs + numBytes - 1)
Friend Sub bytes(pOut() As Byte, ByVal ofs As Long, ByVal numBytes As Long)
    Dim numCopyBytes As Long, i As Long, size As Long
    size = msgData.rawDataSize()
    If index >= size Then
        numCopyBytes = 0
    Else
        numCopyBytes = size - index
        If numBytes < numCopyBytes Then numCopyBytes = numBytes
    End If
    If numCopyBytes > 0 Then
        msgData.copyData pOut, ofs, index, numCopyBytes
    End If
    If numCopyBytes < numBytes Then
        For i = numCopyBytes To numBytes - 1
            pOut(ofs + i) = 0
        Next
    End If
    index = index + numCopyBytes
End Sub

' bytesVector(VecT &vOut, size_t desiredBytes): appended to vOut(0 .. vOutSize - 1)
Friend Sub bytesVector(vOut() As Byte, vOutSize As Long, ByVal desiredBytes As Long)
    Dim numCopyBytes As Long, startIdx As Long, size As Long
    size = msgData.rawDataSize()
    If index >= size Then
        numCopyBytes = 0
    Else
        numCopyBytes = size - index
        If desiredBytes < numCopyBytes Then numCopyBytes = desiredBytes
    End If
    If numCopyBytes > 0 Then
        startIdx = vOutSize
        If startIdx = 0 Then
            ReDim vOut(0 To numCopyBytes - 1)
        Else
            ReDim Preserve vOut(0 To startIdx + numCopyBytes - 1)
        End If
        vOutSize = startIdx + numCopyBytes
        msgData.copyData vOut, startIdx, index, numCopyBytes
    End If
    index = index + numCopyBytes
End Sub

Friend Function valid() As Boolean
    valid = index <= msgData.rawDataSize()
End Function
