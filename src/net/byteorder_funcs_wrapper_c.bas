Attribute VB_Name = "byteorder_funcs_wrapper_c"
' lib/netplay/byteorder_funcs_wrapper.cpp, byteorder_funcs_wrapper.h: Winsock's htonl / htons / ntohl / ntohs (a
' byte swap on Windows' little-endian x86, done here without the call); uint32_t is a Long's bits, uint16_t a
' Long 0..65535. The header's inline unaligned store / load take a Byte array and an offset.
Option Explicit

' uint32_t wz_htonl(uint32_t hostlong)
Public Function wz_htonl(ByVal hostlong As Long) As Long
    wz_htonl = SHL(hostlong And &HFF&, 24) Or SHL(SHRU(hostlong, 8) And &HFF&, 16) Or _
               SHL(SHRU(hostlong, 16) And &HFF&, 8) Or (SHRU(hostlong, 24) And &HFF&)
End Function

' uint16_t wz_htons(uint16_t hostshort)
Public Function wz_htons(ByVal hostshort As Long) As Long
    wz_htons = ((hostshort And &HFF&) * &H100&) Or ((hostshort And &HFF00&) \ &H100&)
End Function

' uint32_t wz_ntohl(uint32_t netlong)
Public Function wz_ntohl(ByVal netlong As Long) As Long
    wz_ntohl = wz_htonl(netlong)
End Function

' uint16_t wz_ntohs(uint16_t netshort)
Public Function wz_ntohs(ByVal netshort As Long) As Long
    wz_ntohs = wz_htons(netshort)
End Function

' inline void wz_htons_store_unaligned(uint8_t* dst, uint16_t src)
Public Sub wz_htons_store_unaligned(dst() As Byte, ByVal ofs As Long, ByVal src As Long)
    dst(ofs) = (src \ &H100&) And &HFF&
    dst(ofs + 1) = src And &HFF&
End Sub

' inline void wz_ntohs_load_unaligned(uint16_t& dst, const uint8_t* src)
Public Sub wz_ntohs_load_unaligned(dst As Long, src() As Byte, ByVal ofs As Long)
    dst = CLng(src(ofs)) * &H100& Or src(ofs + 1)
End Sub
