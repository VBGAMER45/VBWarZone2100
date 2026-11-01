Attribute VB_Name = "ip_helpers_c"
' lib/netplay/ip_helpers.cpp, ip_helpers.h. optional<IPClassification> / optional<IPFormat> are Longs, -1 for
' nullopt.
Option Explicit

Private Declare Function inet_pton Lib "ws2_32" (ByVal Family As Long, ByVal pszAddrString As String, pAddrBuf As Any) As Long

Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23

' enum class IPClassification
Public Const IPClassification_PrivateOrULA As Long = 0   ' IPv4 Private Ranges or IPv6 Unique Local (ULA)
Public Const IPClassification_LinkLocal As Long = 1
Public Const IPClassification_Loopback As Long = 2
Public Const IPClassification_Other As Long = 3

' enum class IPFormat
Public Const IPFormat_IPv4 As Long = 0
Public Const IPFormat_IPv6 As Long = 1
Public Const IPFormat_IPv4InIPv6 As Long = 2              ' an IPv4-mapped IPv6 address, i.e. ::ffff:ip4

' const char* to_string(IPFormat ipFormat)
Public Function IPFormat_to_string(ByVal ipFormat As Long) As String
    Select Case ipFormat
    Case IPFormat_IPv4
        IPFormat_to_string = "ip4"
    Case IPFormat_IPv6
        IPFormat_to_string = "ip6"
    Case IPFormat_IPv4InIPv6
        IPFormat_to_string = "ip4in6"
    Case Else
        IPFormat_to_string = ""                 ' silence gcc warning
    End Select
End Function

' optional<IPClassification> getIPv4Classification(const struct in_addr& addr4): b(0) is the address's first byte
' (the network byte order's high byte)
Private Function getIPv4Classification(b() As Byte, ByVal ofs As Long) As Long
    Dim b1 As Long, b2 As Long
    b1 = b(ofs)
    b2 = b(ofs + 1)
    ' Private IPv4 Ranges (RFC 1918)
    ' 10.0.0.0/8
    If b1 = 10 Then getIPv4Classification = IPClassification_PrivateOrULA: Exit Function
    ' 172.16.0.0/12
    If b1 = 172 And (b2 >= 16 And b2 <= 31) Then getIPv4Classification = IPClassification_PrivateOrULA: Exit Function
    ' 192.168.0.0/16
    If b1 = 192 And b2 = 168 Then getIPv4Classification = IPClassification_PrivateOrULA: Exit Function
    ' Link-Local: 169.254.0.0/16
    If b1 = 169 And b2 = 254 Then getIPv4Classification = IPClassification_LinkLocal: Exit Function
    ' Loopback: 127.0.0.0/8
    If b1 = 127 Then getIPv4Classification = IPClassification_Loopback: Exit Function
    getIPv4Classification = IPClassification_Other
End Function

' (VB) IN6_IS_ADDR_V4MAPPED
Private Function in6IsV4Mapped(a() As Byte) As Boolean
    Dim i As Long
    For i = 0 To 9
        If a(i) <> 0 Then Exit Function
    Next
    in6IsV4Mapped = (a(10) = &HFF And a(11) = &HFF)
End Function

' optional<IPClassification> GetIPClassification(const char *ipAddress) (vbNullString: nullptr)
Public Function GetIPClassification(ByVal ipAddress As String) As Long
    Dim addr6(0 To 15) As Byte, addr4(0 To 3) As Byte, i As Long, loopback As Boolean
    If StrPtr(ipAddress) = 0 Then
        GetIPClassification = -1
        Exit Function
    End If

    If inet_pton(AF_INET6, ipAddress, addr6(0)) > 0 Then
        If Not in6IsV4Mapped(addr6) Then
            ' Process IPv6 Address
            ' Link-local: fe80::/10 (First byte is 0xFE, second byte has high bits 10xx)
            If addr6(0) = &HFE And (addr6(1) And &HC0) = &H80 Then GetIPClassification = IPClassification_LinkLocal: Exit Function
            ' Unique Local (ULA): fc00::/7 (First byte is 0xFC or 0xFD)
            If (addr6(0) And &HFE) = &HFC Then GetIPClassification = IPClassification_PrivateOrULA: Exit Function
            ' Loopback: ::1 (all zeros except last byte is 1)
            loopback = (addr6(15) = 1)
            For i = 0 To 14
                If addr6(i) <> 0 Then loopback = False
            Next
            If loopback Then GetIPClassification = IPClassification_Loopback: Exit Function
            GetIPClassification = IPClassification_Other
        Else
            ' Unmap IPv4 address & fall-through to processing IPv4 address
            GetIPClassification = getIPv4Classification(addr6, 12)
        End If
        Exit Function
    End If

    If inet_pton(AF_INET, ipAddress, addr4(0)) > 0 Then
        GetIPClassification = getIPv4Classification(addr4, 0)
        Exit Function
    End If

    GetIPClassification = -1
End Function

' optional<IPFormat> GetIPFormat(const char *ipAddress) (vbNullString: nullptr)
Public Function GetIPFormat(ByVal ipAddress As String) As Long
    Dim addr6(0 To 15) As Byte, addr4(0 To 3) As Byte
    If StrPtr(ipAddress) = 0 Then
        GetIPFormat = -1
        Exit Function
    End If

    If inet_pton(AF_INET6, ipAddress, addr6(0)) > 0 Then
        If Not in6IsV4Mapped(addr6) Then
            ' Actual IPv6 Address
            GetIPFormat = IPFormat_IPv6
        Else
            ' IPv4-mapped IPv6 address
            GetIPFormat = IPFormat_IPv4InIPv6
        End If
        Exit Function
    End If

    If inet_pton(AF_INET, ipAddress, addr4(0)) > 0 Then
        GetIPFormat = IPFormat_IPv4
        Exit Function
    End If

    GetIPFormat = -1
End Function
