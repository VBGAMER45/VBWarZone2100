VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCPConnectionAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/tcp/tcp_connection_address.cpp, tcp_connection_address.h: class tcp::TCPConnectionAddress, the TCP
' part of an IConnectionAddress: getaddrinfo's list (SocketAddress*: a Long pointer), freed with it.
' This file includes tcp/sock_error.h; EINVAL and EFAULT are not among its defines (MSVC's 22 and 14).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Sub freeaddrinfo Lib "ws2_32" (ByVal pAddrInfo As Long)
Private Declare Function inet_ntop Lib "ws2_32" (ByVal Family As Long, pAddr As Any, ByVal pStringBuf As String, ByVal StringBufSize As Long) As Long

Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23
Private Const EINVAL As Long = 22
Private Const EFAULT As Long = 14

Private addr_ As Long                           ' SocketAddress*

' explicit TCPConnectionAddress(SocketAddress* addr): makes its base, the IConnectionAddress it returns
Friend Function construct(ByVal addr As Long) As IConnectionAddress
    Dim base As IConnectionAddress
    addr_ = addr
    Set base = New IConnectionAddress
    base.construct Me
    Set construct = base
End Function

' virtual ~TCPConnectionAddress() override
Private Sub Class_Terminate()
    If Not (addr_ <> 0) Then ASSERT_FAILURE "addr_ != nullptr", "lib/netplay/tcp/tcp_connection_address.cpp:47", "~TCPConnectionAddress", "Invalid addrinfo stored in the connection address"
    freeaddrinfo addr_
End Sub

' NOTE: The lifetime of the returned `addrinfo` struct is bounded by the parent object's lifetime!
' const SocketAddress* asRawSocketAddress() const
Friend Function asRawSocketAddress() As Long
    asRawSocketAddress = addr_
End Function

' virtual net::result<std::string> toString() const override
Friend Function toString(ec As Long) As String
    Dim ai As ADDRINFOA, sa(0 To 27) As Byte, ipStr As String, r As Long
    ec = 0
    If Not (addr_ <> 0) Then ASSERT_FAILURE "addr_ != nullptr", "lib/netplay/tcp/tcp_connection_address.cpp:53", "toString", "Invalid addrinfo": ec = make_network_error_code(EINVAL): Exit Function
    readAddrinfo addr_, ai
    If ai.ai_family = AF_INET Then
        CopyMemory sa(0), ByVal ai.ai_addr, 16               ' struct sockaddr_in ipv4_addr
        ipStr = String$(16, vbNullChar)                      ' INET_ADDRSTRLEN
        r = inet_ntop(ai.ai_family, sa(4), ipStr, 16)
        sockCalled
        If r <> 0 Then
            ' fmt::format("{}:{}", ipStr, ntohs(ipv4_addr.sin_port))
            toString = Left$(ipStr, InStr(ipStr, vbNullChar) - 1) & ":" & CStr(CLng(sa(2)) * 256 + sa(3))
            Exit Function
        End If
        ec = make_network_error_code(getSockErr())
        Exit Function
    ElseIf ai.ai_family = AF_INET6 Then
        CopyMemory sa(0), ByVal ai.ai_addr, 28               ' struct sockaddr_in6 ipv6_addr
        ipStr = String$(46, vbNullChar)                      ' INET6_ADDRSTRLEN
        r = inet_ntop(ai.ai_family, sa(8), ipStr, 46)
        sockCalled
        If r <> 0 Then
            ' fmt::format("[{}]:{}", ipStr, ntohs(ipv6_addr.sin6_port))
            toString = "[" & Left$(ipStr, InStr(ipStr, vbNullChar) - 1) & "]:" & CStr(CLng(sa(2)) * 256 + sa(3))
            Exit Function
        End If
        ec = make_network_error_code(getSockErr())
        Exit Function
    End If
    ec = make_network_error_code(EFAULT)        ' Unknown family
End Function
