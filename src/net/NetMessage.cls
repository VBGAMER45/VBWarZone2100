VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/netplay/netqueue.h, netqueue.cpp: class NetMessage, a network message
' with a type (uint8_t) and payload, the meaning of which depends on the
' type. Instances are immutable after creation and are made by
' NetMessageBuilder (build), so where C++ copies one the port shares the
' object. The binary layout (data_):
'
' | Type (1 byte) | Payload Length (2 bytes) | Payload (variable length) |
'
' The payload length is a uint16_t in network byte order (big endian).
' NetMessage::tryFromRawData is netqueue_c's NetMessage_tryFromRawData.
Option Explicit

Private data_() As Byte                 ' NetMsgDataVector
Private dataSize As Long                ' (VB) data_.size()

' Meant to be executed only by NetMessageBuilder: takes data(0 .. size - 1)
' (the vector moved in)
Friend Sub construct(data() As Byte, ByVal size As Long)
    data_ = data
    dataSize = size
End Sub

' type()
Public Function type_() As Long
    If Not (dataSize > 0) Then ASSERT_FAILURE "!data_.empty()", "lib/netplay/netqueue.cpp:116", "type", "Invalid message data": type_ = 0: Exit Function
    type_ = data_(0)
End Function

' rawData(): a copy of the bytes (C++ returns a const reference)
Public Function rawData() As Byte()
    Dim b() As Byte
    If dataSize > 0 Then
        b = data_
        If UBound(b) <> dataSize - 1 Then ReDim Preserve b(0 To dataSize - 1)
    End If
    rawData = b
End Function

' (VB) rawData().size(), without the copy
Friend Function rawDataSize() As Long
    rawDataSize = dataSize
End Function

' Returns a copy of the payload data (offset by HEADER_LENGTH from the beginning of the raw data); empty for nullptr.
Public Function payload() As Byte()
    Dim b() As Byte, i As Long
    If Not (dataSize > 0 And dataSize >= NetMessage_HEADER_LENGTH) Then ASSERT_FAILURE "!data_.empty() && data_.size() >= HEADER_LENGTH", "lib/netplay/netqueue.cpp:127", "payload", "Invalid message data": payload = b: Exit Function
    If dataSize > NetMessage_HEADER_LENGTH Then
        ReDim b(0 To dataSize - NetMessage_HEADER_LENGTH - 1)
        For i = 0 To dataSize - NetMessage_HEADER_LENGTH - 1
            b(i) = data_(NetMessage_HEADER_LENGTH + i)
        Next
    End If
    payload = b
End Function

Public Function payloadSize() As Long
    If Not (dataSize >= NetMessage_HEADER_LENGTH) Then ASSERT_FAILURE "data_.size() >= HEADER_LENGTH", "lib/netplay/netqueue.cpp:133", "payloadSize", "Invalid message data": payloadSize = 0: Exit Function
    payloadSize = dataSize - NetMessage_HEADER_LENGTH
End Function

' Append the raw data of this message to the provided output vector
' (output(0 .. outputSize - 1), grown as needed).
Friend Sub rawDataAppendToVector(output() As Byte, outputSize As Long)
    Dim oldLen As Long, i As Long
    oldLen = outputSize
    If dataSize = 0 Then Exit Sub
    If oldLen = 0 Then
        ReDim output(0 To oldLen + dataSize - 1)
    ElseIf UBound(output) < oldLen + dataSize - 1 Then
        ReDim Preserve output(0 To oldLen + dataSize - 1)
    End If
    For i = 0 To dataSize - 1
        output(oldLen + i) = data_(i)
    Next
    outputSize = oldLen + dataSize
End Sub

' (VB) (*msgData)[i], for MessageReader
Friend Function dataByte(ByVal i As Long) As Byte
    dataByte = data_(i)
End Function

' (VB) memcpy(pOut + outOfs, &data_[srcOfs], n), for MessageReader
Friend Sub copyData(pOut() As Byte, ByVal outOfs As Long, ByVal srcOfs As Long, ByVal n As Long)
    Dim i As Long
    For i = 0 To n - 1
        pOut(outOfs + i) = data_(srcOfs + i)
    Next
End Sub
