Attribute VB_Name = "netlog_c"
' lib/netplay/netlog.cpp, netlog.h: the packet counters (NETlogPacket).
' The net log file (NETstartLogging, NETstopLogging, NETlogEntry's
' writing) comes with the connections (cluster J): NETstartLogging is
' never called yet, so there is no log file and NETlogEntry returns false
' as it does without one.
Option Explicit

Private Const NUM_GAME_PACKETS As Long = 256

Private packetcount(0 To 1, 0 To NUM_GAME_PACKETS - 1) As Long  ' uint32_t
Private packetsize(0 To 1, 0 To NUM_GAME_PACKETS - 1) As Long   ' uint32_t

' log packet
'  type, uint8_t, the packet's type.
'  size, uint32_t, the packet's size
'  received, bool, true if we are receiving a packet, false if we are sending a packet.
Public Sub NETlogPacket(ByVal typ As Long, ByVal size As Long, ByVal received As Boolean)
    Dim r As Long
    ' STATIC_ASSERT((1 << (8 * sizeof(type))) == NUM_GAME_PACKETS); // NUM_GAME_PACKETS must be larger than maximum possible type.
    If received Then r = 1 Else r = 0
    typ = typ And &HFF&
    packetcount(r, typ) = FromU32(U32(packetcount(r, typ)) + 1)
    packetsize(r, typ) = FromU32(U32(packetsize(r, typ)) + U32(size))
End Sub

' NETlogEntry(str, a, b): no log file is open (pFileHandle is nullptr)
Public Function NETlogEntry(ByVal str_ As String, ByVal a As Long, ByVal b As Long) As Boolean
    NETlogEntry = False
End Function
