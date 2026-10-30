Attribute VB_Name = "netqueue_c"
' lib/netplay/netqueue.cpp, lib/netplay/netqueue.h: the variable-length uint32_t
' encoding, NetMessage::tryFromRawData and the classes' factories. The
' classes are NetMessage (public: NETgetMessage returns one),
' NetMessageBuilder, NetQueue and NetQueuePair; MessageWriter and
' MessageReader (netqueue.h) are public classes too.
Option Explicit

' NetMessage::HEADER_LENGTH: (type) + (payload length)
Public Const NetMessage_HEADER_LENGTH As Long = 3

' See comments in netqueue.h.


' Byte n is the final byte, iff it is less than 256-a[n].

' Some possible values of table_uint32_t_a[n], such that the maximum encodable value is 0xFFFFFFFF, satisfying the assertion below:
' ASSERT(0xFFFFFFFF == 255*(1 + a[0]*(1 + a[1]*(1 + a[2]*(1 + a[3]*(1 + a[4]))))), "Maximum encodable value not 0xFFFFFFFF.");

' constexpr unsigned table_uint32_t_a[5] = {78, 95, 32, 70, 0}, table_uint32_t_m[5] = {1, 78, 7410, 237120, 16598400};
' <178: 1 byte, <12736: 2 bytes, <1672576: 3 bytes, <45776896: 4 bytes, <4294967296: 5 bytes
Private Function table_uint32_t_a(ByVal n As Long) As Long
    Select Case n
    Case 0: table_uint32_t_a = 78
    Case 1: table_uint32_t_a = 95
    Case 2: table_uint32_t_a = 32
    Case 3: table_uint32_t_a = 70
    Case Else: table_uint32_t_a = 0
    End Select
End Function

Private Function table_uint32_t_m(ByVal n As Long) As Long
    Select Case n
    Case 0: table_uint32_t_m = 1
    Case 1: table_uint32_t_m = 78
    Case 2: table_uint32_t_m = 7410
    Case 3: table_uint32_t_m = 237120
    Case Else: table_uint32_t_m = 16598400
    End Select
End Function

' Returns the number of bytes required to encode v.
' (VB) the unsigned arithmetic in Double
Public Function encodedlength_uint32_t(ByVal v As Long) As Long
    Dim n As Long, a As Long, m As Long, possibleValues As Double, vd As Double
    vd = U32(v)
    Do
        a = table_uint32_t_a(n)
        m = table_uint32_t_m(n)
        possibleValues = CDbl(256 - a) * m      ' Number of values which encode to n + 1 bytes.
        If vd < possibleValues Then
            encodedlength_uint32_t = n + 1
            Exit Function
        End If
        vd = vd - possibleValues
        n = n + 1
    Loop
End Function

' Returns true iff there is another byte to be encoded.
' Modifies v.
' Input is v, output is b.
Public Function encode_uint32_t(b As Long, v As Long, ByVal n As Long) As Boolean
    Dim a As Long, isLastByte As Boolean, vd As Double, q As Double
    a = table_uint32_t_a(n)

    isLastByte = U32(v) < 256 - a
    If isLastByte Then
        b = v And &HFF&
    Else
        vd = U32(v) - (256 - a)
        q = Int(vd / a)
        b = 255 - (vd - q * a)
        v = FromU32(q)
    End If
    encode_uint32_t = Not isLastByte
End Function

' Returns true iff another byte is required to decode.
' Must init v to 0, does not modify b.
' Input is b, output is v.
Public Function decode_uint32_t(ByVal b As Long, v As Long, ByVal n As Long) As Boolean
    Dim a As Long, m As Long, isLastByte As Boolean
    a = table_uint32_t_a(n)
    m = table_uint32_t_m(n)

    isLastByte = b < 256 - a
    If isLastByte Then
        v = FromU32(U32(v) + CDbl(b) * m)
    Else
        v = FromU32(U32(v) + CDbl(256 - a + 255 - b) * m)
    End If
    decode_uint32_t = Not isLastByte
End Function

' NetMessage::tryFromRawData(buffer, bufferLen): Nothing for nullopt
Public Function NetMessage_tryFromRawData(buffer() As Byte, ByVal bufferLen As Long) As NetMessage
    Dim typ As Long, length As Long, msg As NetMessageBuilder
    If bufferLen < NetMessage_HEADER_LENGTH Then Exit Function

    typ = buffer(0)
    ' Load payload length from uint16_t (network byte order) starting at the second byte of the data buffer.
    length = CLng(buffer(1)) * 256 + buffer(2)

    If bufferLen - NetMessage_HEADER_LENGTH < length Then
        Exit Function                           ' Don't have a whole message ready yet.
    End If

    Set msg = NetMessageBuilder_new(typ, length)
    msg.append_bytes buffer, NetMessage_HEADER_LENGTH, length
    Set NetMessage_tryFromRawData = msg.build()
End Function

' NetMessageBuilder(uint8_t type, size_t reservedCapacity = 16)
Public Function NetMessageBuilder_new(ByVal typ As Long, Optional ByVal reservedCapacity As Long = 16) As NetMessageBuilder
    Set NetMessageBuilder_new = New NetMessageBuilder
    NetMessageBuilder_new.construct typ, reservedCapacity
End Function

' NetMessageBuilder(NetMsgDataVector&& rawData): rawData(0 .. rawDataSize - 1)
Public Function NetMessageBuilder_new_raw(rawData() As Byte, ByVal rawDataSize As Long) As NetMessageBuilder
    Set NetMessageBuilder_new_raw = New NetMessageBuilder
    NetMessageBuilder_new_raw.construct_raw rawData, rawDataSize
End Function

' MessageWriter(NETQUEUE queue, uint8_t messageType)
Public Function MessageWriter_new(ByVal queue As Long, ByVal messageType As Long) As MessageWriter
    Set MessageWriter_new = New MessageWriter
    MessageWriter_new.construct queue, messageType
End Function

' MessageReader(const NetMessage& m)
Public Function MessageReader_new(ByVal m As NetMessage) As MessageReader
    Set MessageReader_new = New MessageReader
    MessageReader_new.construct m
End Function
