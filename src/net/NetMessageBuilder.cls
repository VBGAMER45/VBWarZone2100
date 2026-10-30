VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetMessageBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/netqueue.h, netqueue.cpp: class NetMessageBuilder, which
' builds a network message: append() adds data, build() moves the data into
' the NetMessage being built (after which this builder is invalid). Made by
' netqueue_c's NetMessageBuilder_new (type, reserved capacity) and
' NetMessageBuilder_new_raw (the raw data moved in).
Option Explicit

Private data_() As Byte                 ' NetMsgDataVector
Private dataSize As Long                ' (VB) data_.size()
Private capacity As Long                ' (VB) data_'s allocated size (0: not allocated)

' (VB) data_.reserve(n)
Private Sub reserve(ByVal n As Long)
    Dim cap As Long
    If capacity >= n Then Exit Sub
    cap = capacity * 2
    If cap < n Then cap = n
    If capacity = 0 Then
        ReDim data_(0 To cap - 1)
    Else
        ReDim Preserve data_(0 To cap - 1)
    End If
    capacity = cap
End Sub

' NetMessageBuilder(uint8_t type, size_t reservedCapacity = 16)
Friend Sub construct(ByVal typ As Long, ByVal reservedCapacity As Long)
    reserve reservedCapacity + NetMessage_HEADER_LENGTH
    dataSize = NetMessage_HEADER_LENGTH         ' data_.resize(NetMessage::HEADER_LENGTH)
    data_(0) = typ And &HFF&
End Sub

' NetMessageBuilder(NetMsgDataVector&& rawData): rawData(0 .. rawDataSize - 1)
Friend Sub construct_raw(rawData() As Byte, ByVal rawDataSize As Long)
    If rawDataSize > 0 Then
        data_ = rawData
        capacity = UBound(data_) + 1
    End If
    dataSize = rawDataSize
End Sub

' type()
Friend Function type_() As Long
    type_ = data_(0)
End Function

' append(uint8_t v)
Friend Sub append(ByVal v As Long)
    reserve dataSize + 1
    data_(dataSize) = v And &HFF&
    dataSize = dataSize + 1
End Sub

' append(const uint8_t *src, size_t len): src(srcOfs .. srcOfs + length - 1)
Friend Sub append_bytes(src() As Byte, ByVal srcOfs As Long, ByVal length As Long)
    Dim resultLen As Long, i As Long
    resultLen = dataSize + length
    If Not (resultLen <= 65535) Then ASSERT_FAILURE "resultLen <= UINT16_MAX", "lib/netplay/netqueue.h:126", "append", "Resulting message length exceeds uint16_t max: %zu", resultLen: Exit Sub
    If length <= 0 Then Exit Sub
    reserve resultLen
    For i = 0 To length - 1
        data_(dataSize + i) = src(srcOfs + i)
    Next
    dataSize = resultLen
End Sub

' Build the final message (invalidates NetMessageBuilder instance)
Friend Function build() As NetMessage
    Dim length As Long, payloadLen As Long
    ' (VB) a raw vector shorter than the header: C++ writes past its end; here it is padded
    If dataSize < NetMessage_HEADER_LENGTH Then
        reserve NetMessage_HEADER_LENGTH
        dataSize = NetMessage_HEADER_LENGTH
    End If
    ' Update the length in the header
    length = dataSize - NetMessage_HEADER_LENGTH
    If length > 65535 Then
        ASSERT_FAILURE "false", "lib/netplay/netqueue.h:137", "build", "Message payload length exceeds uint16_t max: %zu (message type: %u)", length, type_()
        length = 65535                          ' Clamp to max length, so we can still send the message and probably gracefully handle it on the other side.
    End If
    payloadLen = length
    ' Store payload length in network byte order starting at the second byte of the data vector.
    data_(1) = (payloadLen \ 256) And &HFF&
    data_(2) = payloadLen And &HFF&
    Set build = New NetMessage
    build.construct data_, dataSize
    ' (the data moved out)
    Erase data_
    dataSize = 0
    capacity = 0
End Function
