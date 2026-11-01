Attribute VB_Name = "plum_net_c"
' libplum: Copyright (c) 2020-2021 Paul-Louis Ageneau; 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/net.c, net.h (the _WIN32 branch: GetIpForwardTable2).
' (VB) TEST ONLY: the simulated router override (tools/portmap_test.py). When the environment variable
' WZ_TEST_PORTMAPPING_GATEWAY is set, the default gateway is that address instead of the routing table's, so PCP and
' NAT-PMP go to <it>:5351, and UPnP's SSDP search goes unicast to <it>:<WZ_TEST_PORTMAPPING_SSDP_PORT, default 1900>
' instead of the multicast group (net_test_override_ssdp, called by plum_upnp_c's upnp_impl_probe). Only a loopback
' IPv4 address (127.x.x.x) is taken: any other value makes gateway discovery and the SSDP search fail, so nothing a
' test sends can reach the real router or the LAN. Nothing reads these variables but these two functions.
Option Explicit

Private Declare Function ws_socket Lib "ws2_32" Alias "socket" (ByVal af As Long, ByVal type_ As Long, ByVal protocol As Long) As Long
Private Declare Function ws_connect Lib "ws2_32" Alias "connect" (ByVal s As Long, name_ As Any, ByVal namelen As Long) As Long
Private Declare Function ws_getsockname Lib "ws2_32" Alias "getsockname" (ByVal s As Long, name_ As Any, namelen As Long) As Long
Private Declare Function ws_closesocket Lib "ws2_32" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function GetIpForwardTable2 Lib "iphlpapi" (ByVal Family As Long, Table As Long) As Long
Private Declare Sub FreeMibTable Lib "iphlpapi" (ByVal Memory As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Const INVALID_SOCKET As Long = -1
Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23
Private Const SOCK_DGRAM As Long = 2
Private Const IPPROTO_UDP As Long = 17
Private Const NO_ERROR As Long = 0
' netioapi.h: MIB_IPFORWARD_TABLE2 { ULONG NumEntries; MIB_IPFORWARD_ROW2 Table[] (at 8) }, a row's size and fields
Private Const MIB_IPFORWARD_ROW2_SIZE As Long = 104
Private Const ROW_PREFIX_LENGTH As Long = 40    ' DestinationPrefix.PrefixLength
Private Const ROW_NEXT_HOP As Long = 44         ' NextHop (SOCKADDR_INET): si_family, Ipv4.sin_addr at +4, Ipv6.sin6_addr at +8

Private Const UPNP_SSDP_PORT As Long = 1900     ' upnp.h

Private Function socket_(ByVal af As Long, ByVal type_ As Long, ByVal protocol As Long) As Long
    socket_ = ws_socket(af, type_, protocol)
    plum_sock_called
End Function

Private Function connect_(ByVal s As Long, name_() As Byte, ByVal namelen As Long) As Long
    connect_ = ws_connect(s, name_(0), namelen)
    plum_sock_called
End Function

Private Function getsockname(ByVal s As Long, name_() As Byte, namelen As Long) As Long
    getsockname = ws_getsockname(s, name_(0), namelen)
    plum_sock_called
End Function

Private Function closesocket(ByVal s As Long) As Long
    closesocket = ws_closesocket(s)
    plum_sock_called
End Function

' int net_get_default_interface(int family, addr_record_t *record)
Public Function net_get_default_interface(ByVal family As Long, record As addr_record_t) As Long
    Dim host As String, dummy As addr_record_t, sock As Long
    Select Case family
    Case AF_INET
        host = "192.0.2.1"      ' dummy public unreachable address
    Case AF_INET6
        host = "2001:db8::1"    ' dummy public unreachable address
    Case Else
        plum_log_write PLUM_LOG_LEVEL_WARN, "Unknown address family %d", family
        net_get_default_interface = -1
        Exit Function
    End Select

    addr_set family, host, 9, dummy

    sock = socket_(family, SOCK_DGRAM, IPPROTO_UDP)
    If sock = INVALID_SOCKET Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "UDP socket creation failed, errno=%d", sockerrno()
        net_get_default_interface = -1
        Exit Function
    End If

    If connect_(sock, dummy.addr, dummy.len_) <> 0 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "connect failed on UDP socket, errno=%d", sockerrno()
        GoTo error_
    End If

    record.len_ = 128                           ' sizeof(record->addr)
    If getsockname(sock, record.addr, record.len_) <> 0 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "getsockname failed, errno=%d", sockerrno()
        GoTo error_
    End If

    addr_unmap_inet6_v4mapped record.addr, record.len_
    If sa_family(record.addr) <> family Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "getsockname returned unexpected address family"
        GoTo error_
    End If

    addr_set_port record.addr, 0

    If addr_is_local(record.addr) Then GoTo error_

    closesocket sock
    net_get_default_interface = 0
    Exit Function

error_:
    closesocket sock
    net_get_default_interface = -1
End Function

' int net_get_default_gateway(int family, addr_record_t *record)
' (VB) the _WIN32 branch is net_get_default_gateway_os; this applies the test-only override after it
Public Function net_get_default_gateway(ByVal family As Long, record As addr_record_t) As Long
    Dim ret As Long, overrideRet As Long
    ret = net_get_default_gateway_os(family, record)
    overrideRet = net_test_override_gateway(family, record)
    If overrideRet <> 0 Then ret = IIf(overrideRet > 0, 0, -1)
    net_get_default_gateway = ret
End Function

' the _WIN32 branch of net_get_default_gateway
Private Function net_get_default_gateway_os(ByVal family As Long, record As addr_record_t) As Long
    Dim table As Long, numEntries As Long, n As Long, row As Long, prefixLength As Byte, siFamily As Integer
    Dim b(0 To 15) As Byte
    If GetIpForwardTable2(family, table) <> NO_ERROR Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "GetIpForwardTable2 failed"
        net_get_default_gateway_os = -1
        Exit Function
    End If

    CopyMemory numEntries, ByVal table, 4
    For n = 0 To numEntries - 1
        row = table + 8 + n * MIB_IPFORWARD_ROW2_SIZE   ' MIB_IPFORWARD_ROW2 *row = table->Table + n;
        CopyMemory prefixLength, ByVal row + ROW_PREFIX_LENGTH, 1
        If prefixLength = 0 Then
            CopyMemory siFamily, ByVal row + ROW_NEXT_HOP, 2   ' SOCKADDR_INET *nextHop = &row->NextHop;
            Select Case siFamily
            Case AF_INET
                CopyMemory b(0), ByVal row + ROW_NEXT_HOP + 4, 4
                addr_set_binary AF_INET, b, 0, 0, record
                net_get_default_gateway_os = 0
                Exit Function
            Case AF_INET6
                CopyMemory b(0), ByVal row + ROW_NEXT_HOP + 8, 16
                addr_set_binary AF_INET6, b, 0, 0, record
                net_get_default_gateway_os = 0
                Exit Function
            Case Else
                plum_log_write PLUM_LOG_LEVEL_WARN, "GetIpForwardTable2 returned unexpected address family"
            End Select
        End If
    Next

    plum_log_write PLUM_LOG_LEVEL_WARN, "No default route found"
    FreeMibTable table
    net_get_default_gateway_os = -1
End Function

' (VB) TEST ONLY (see the module's comment): the test override's IPv4 loopback address; 0 when the variable is not
' set, 1 when record holds it, -1 when the variable holds anything but a loopback IPv4 address
Private Function net_test_override_address(record As addr_record_t) As Long
    Dim value As String
    value = Environ$("WZ_TEST_PORTMAPPING_GATEWAY")
    If value = "" Then Exit Function
    If addr_set(AF_INET, value, 0, record) = 0 Then
        If record.addr(4) = 127 Then
            net_test_override_address = 1
            Exit Function
        End If
    End If
    plum_log_write PLUM_LOG_LEVEL_ERROR, "(test override) WZ_TEST_PORTMAPPING_GATEWAY=%s is not a loopback IPv4 address: refused", value
    net_test_override_address = -1
End Function

' (VB) TEST ONLY: the default gateway replaced by the override; 0 when not set, 1 when applied, -1 when refused
Private Function net_test_override_gateway(ByVal family As Long, record As addr_record_t) As Long
    Dim test As addr_record_t, found As String, ret As Long
    ret = net_test_override_address(test)
    If ret = 0 Then Exit Function
    If ret > 0 And family = AF_INET Then
        If record.len_ > 0 Then addr_record_to_string record, found, ADDR_MAX_STRING_LEN Else found = "none"
        record = test
        plum_log_write PLUM_LOG_LEVEL_INFO, "(test override) default gateway %s replaced by %s", found, Environ$("WZ_TEST_PORTMAPPING_GATEWAY")
        net_test_override_gateway = 1
    Else
        net_test_override_gateway = -1
    End If
End Function

' (VB) TEST ONLY: called by upnp_impl_probe with the SSDP multicast address; the override's address and SSDP port
' instead. 0 when the variable is not set, 1 when applied, -1 when refused (the caller must not send)
Public Function net_test_override_ssdp(broadcast As addr_record_t) As Long
    Dim test As addr_record_t, ret As Long, port As Long
    ret = net_test_override_address(test)
    If ret <= 0 Then
        net_test_override_ssdp = ret
        Exit Function
    End If
    port = Val(Environ$("WZ_TEST_PORTMAPPING_SSDP_PORT"))
    If port <= 0 Or port > 65535 Then port = UPNP_SSDP_PORT
    addr_set_port test.addr, port
    broadcast = test
    plum_log_write PLUM_LOG_LEVEL_INFO, "(test override) SSDP search sent unicast to %s:%d", Environ$("WZ_TEST_PORTMAPPING_GATEWAY"), port
    net_test_override_ssdp = 1
End Function
