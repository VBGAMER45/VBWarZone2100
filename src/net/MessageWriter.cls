VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MessageWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/netplay/netqueue.h: class MessageWriter, used for serialising, using
' the same interface as MessageReader. Only NETbeginEncode makes one
' (netqueue_c's MessageWriter_new); the exe holds it and passes it to
' NetApi's NET* procedures. Its members are for WzNet only (Friend): the
' queue (a NETQUEUE Long), the secret-wrap flag and the message builder.
Option Explicit

Private queueInfo_ As Long              ' NETQUEUE queueInfo
Private bSecretMessageWrap_ As Boolean
Private msgBuilder_ As NetMessageBuilder

' This should allocate a message owned by this instance
Friend Sub construct(ByVal queue As Long, ByVal messageType As Long)
    queueInfo_ = queue
    Set msgBuilder_ = NetMessageBuilder_new(messageType)
End Sub

Friend Property Get queueInfo() As Long
    queueInfo = queueInfo_
End Property

Friend Property Get bSecretMessageWrap() As Boolean
    bSecretMessageWrap = bSecretMessageWrap_
End Property

Friend Property Let bSecretMessageWrap(ByVal value As Boolean)
    bSecretMessageWrap_ = value
End Property

Friend Property Get msgBuilder() As NetMessageBuilder
    Set msgBuilder = msgBuilder_
End Property

' byte(uint8_t v)
Friend Sub byte_(ByVal v As Long)
    msgBuilder_.append v
End Sub

' bytes(const uint8_t *pIn, size_t numBytes): pIn(ofs .. ofs + numBytes - 1)
Friend Sub bytes(pIn() As Byte, ByVal ofs As Long, ByVal numBytes As Long)
    msgBuilder_.append_bytes pIn, ofs, numBytes
End Sub

' bytesVector(const std::vector<uint8_t> &vIn, size_t numBytes): vIn(0 .. vInSize - 1)
Friend Sub bytesVector(vIn() As Byte, ByVal vInSize As Long, ByVal numBytes As Long)
    If numBytes > vInSize Then numBytes = vInSize
    msgBuilder_.append_bytes vIn, 0, numBytes
End Sub

Friend Function valid() As Boolean
    valid = True
End Function
