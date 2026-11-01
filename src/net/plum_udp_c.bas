Attribute VB_Name = "plum_udp_c"
' libplum: Copyright (c) 2020-2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/udp.c, udp.h (the _WIN32 branches: NO_PMTUDISC, not __linux__). A datagram's data is a Byte array
' (from its first byte). udp_set_diffserv (used by no caller) is left out.
Option Explicit

Private Declare Function ws_socket Lib "ws2_32" Alias "socket" (ByVal af As Long, ByVal type_ As Long, ByVal protocol As Long) As Long
Private Declare Function ws_closesocket Lib "ws2_32" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function ws_bind Lib "ws2_32" Alias "bind" (ByVal s As Long, ByVal name_ As Long, ByVal namelen As Long) As Long
Private Declare Function ws_ioctlsocket Lib "ws2_32" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function ws_setsockopt Lib "ws2_32" Alias "setsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function ws_getsockname Lib "ws2_32" Alias "getsockname" (ByVal s As Long, name_ As Any, namelen As Long) As Long
Private Declare Function ws_recvfrom Lib "ws2_32" Alias "recvfrom" (ByVal s As Long, buf As Any, ByVal len_ As Long, ByVal flags As Long, from As Any, fromlen As Long) As Long
Private Declare Function ws_sendto Lib "ws2_32" Alias "sendto" (ByVal s As Long, buf As Any, ByVal len_ As Long, ByVal flags As Long, to_ As Any, ByVal tolen As Long) As Long
Private Declare Function ws_getaddrinfo Lib "ws2_32" Alias "getaddrinfo" (ByVal pNodeName As String, ByVal pServiceName As String, pHints As ADDRINFOA, ppResult As Long) As Long
Private Declare Sub ws_freeaddrinfo Lib "ws2_32" Alias "freeaddrinfo" (ByVal pAddrInfo As Long)
Private Declare Function ws_inet_pton Lib "ws2_32" Alias "inet_pton" (ByVal Family As Long, ByVal pszAddrString As String, pAddrBuf As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Const INVALID_SOCKET As Long = -1
Private Const AF_UNSPEC As Long = 0
Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23
Private Const SOCK_DGRAM As Long = 2
Private Const IPPROTO_IP As Long = 0
Private Const IPPROTO_UDP As Long = 17
Private Const IPPROTO_IPV6 As Long = 41
Private Const IPV6_V6ONLY_ As Long = 27         ' IPV6_V6ONLY
Private Const SOL_SOCKET As Long = &HFFFF&
Private Const SO_BROADCAST As Long = &H20
Private Const SO_REUSEADDR As Long = 4
Private Const IP_ADD_MEMBERSHIP As Long = 12
Private Const IP_DONTFRAGMENT As Long = 14      ' IP_DONTFRAG
Private Const FIONBIO As Long = &H8004667E
Private Const AI_PASSIVE As Long = &H1
Private Const AI_NUMERICSERV As Long = &H8

' typedef struct udp_socket_config
Public Type udp_socket_config_t
    family As Long                              ' AF_UNSPEC means any
    port As Long                                ' 0 means any
    bind_address As String                      ' vbNullString (NULL) means any
    multicast_group As String                   ' vbNullString: none
    enable_broadcast As Boolean
    enable_reuseaddr As Boolean
    enable_dontfrag As Boolean
End Type

Private Function closesocket(ByVal s As Long) As Long
    closesocket = ws_closesocket(s)
    plum_sock_called
End Function

Private Function setsockopt(ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval() As Byte, ByVal optlen As Long) As Long
    setsockopt = ws_setsockopt(s, level, optname, optval(0), optlen)
    plum_sock_called
End Function

Private Function setsockopt_int(ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByVal optval As Long) As Long
    setsockopt_int = ws_setsockopt(s, level, optname, optval, 4)    ' (const char *)&optval, sizeof(sockopt_t)
    plum_sock_called
End Function

Private Function getsockname(ByVal s As Long, name_() As Byte, namelen As Long) As Long
    getsockname = ws_getsockname(s, name_(0), namelen)
    plum_sock_called
End Function

' (VB) a datagram's data from its first byte; size 0 passes a valid pointer all the same
Private Function sendto_(ByVal s As Long, data() As Byte, ByVal size As Long, to_() As Byte, ByVal tolen As Long) As Long
    sendto_ = ws_sendto(s, data(0), size, 0, to_(0), tolen)
    plum_sock_called
End Function

' (VB) inet_pton(family, host, buf + offset)
Private Function inet_pton(ByVal family As Long, ByVal host As String, buf() As Byte, ByVal offset As Long) As Long
    inet_pton = ws_inet_pton(family, host, buf(offset))
    plum_sock_called
End Function

' static struct addrinfo *find_family(struct addrinfo *ai_list, int family)
Private Function find_family(ByVal ai_list As Long, ByVal family As Long) As Long
    Dim ai As Long, info As ADDRINFOA
    ai = ai_list
    Do While ai <> 0
        CopyMemory info, ByVal ai, LenB(info)
        If info.ai_family = family Then Exit Do
        ai = info.ai_next
    Loop
    find_family = ai
End Function

' socket_t udp_create_socket(const udp_socket_config_t *config)
Public Function udp_create_socket(config As udp_socket_config_t) As Long
    Dim sock As Long, hints As ADDRINFOA, bind_service As String, ai_list As Long, ai As Long, info As ADDRINFOA
    Dim mreq(0 To 7) As Byte, nbio As Long, sin_addr(0 To 3) As Byte
    sock = INVALID_SOCKET

    If config.family <> AF_UNSPEC And config.family <> AF_INET And config.family <> AF_INET6 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Unknown address family for UDP socket"
        udp_create_socket = INVALID_SOCKET
        Exit Function
    End If

    ' Obtain local Address
    hints.ai_family = config.family
    hints.ai_socktype = SOCK_DGRAM
    hints.ai_protocol = IPPROTO_UDP
    hints.ai_flags = AI_PASSIVE Or AI_NUMERICSERV
    bind_service = CStr(config.port)            ' snprintf(bind_service, 8, "%hu", config->port)

    ai_list = 0
    If ws_getaddrinfo(config.bind_address, bind_service, hints, ai_list) <> 0 Then
        plum_sock_called
        plum_log_write PLUM_LOG_LEVEL_ERROR, "getaddrinfo for binding address failed, errno=%d", sockerrno()
        udp_create_socket = INVALID_SOCKET
        Exit Function
    End If

    If config.family = AF_UNSPEC Then
        ' Prefer IPv6
        ai = find_family(ai_list, AF_INET6)
        If ai = 0 Then ai = find_family(ai_list, AF_INET)
        If ai = 0 Then
            plum_log_write PLUM_LOG_LEVEL_ERROR, "getaddrinfo for binding address failed: no suitable address family"
            GoTo error_
        End If
    Else
        ai = ai_list
    End If
    CopyMemory info, ByVal ai, LenB(info)

    ' Create socket
    sock = ws_socket(info.ai_family, info.ai_socktype, info.ai_protocol)
    plum_sock_called
    If sock = INVALID_SOCKET Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "UDP socket creation failed, errno=%d", sockerrno()
        GoTo error_
    End If

    ' const sockopt_t enabled = 1; const sockopt_t disabled = 0;

    ' Listen on both IPv6 and IPv4
    If config.family = AF_UNSPEC And info.ai_family = AF_INET6 Then setsockopt_int sock, IPPROTO_IPV6, IPV6_V6ONLY_, 0

    If config.enable_broadcast Then setsockopt_int sock, SOL_SOCKET, SO_BROADCAST, 1

    If config.enable_reuseaddr Then setsockopt_int sock, SOL_SOCKET, SO_REUSEADDR, 1

    If config.enable_dontfrag Then
        ' Set DF flag (NO_PMTUDISC: IP_DONTFRAG is IP_DONTFRAGMENT; no IPV6_DONTFRAG)
        setsockopt_int sock, IPPROTO_IP, IP_DONTFRAGMENT, 1
    End If

    If config.multicast_group <> vbNullString Then
        ' struct ip_mreq mreq (imr_multiaddr at 0, imr_interface at 4)
        If inet_pton(AF_INET, config.multicast_group, mreq, 0) <> 1 Then
            plum_log_write PLUM_LOG_LEVEL_ERROR, "Invalid multicast group"
            GoTo error_
        End If

        If info.ai_family = AF_INET Then
            CopyMemory sin_addr(0), ByVal info.ai_addr + 4, 4   ' ((const struct sockaddr_in *)ai->ai_addr)->sin_addr.s_addr
            CopyMemory mreq(4), sin_addr(0), 4
        Else
            ' htonl(INADDR_ANY): 0
        End If

        setsockopt sock, IPPROTO_IP, IP_ADD_MEMBERSHIP, mreq, 8

        ' TODO
        ' struct ipv6_mreq mreq;
        ' setsockopt(sock, IPPROTO_IPV6, IPV6_JOIN_GROUP, &mreq, sizeof(mreq))
    End If

    nbio = 1
    If ws_ioctlsocket(sock, FIONBIO, nbio) <> 0 Then
        plum_sock_called
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Setting non-blocking mode on UDP socket failed, errno=%d", sockerrno()
        GoTo error_
    End If

    ' Bind it
    If ws_bind(sock, info.ai_addr, info.ai_addrlen) <> 0 Then
        plum_sock_called
        plum_log_write PLUM_LOG_LEVEL_ERROR, "UDP socket binding failed, errno=%d", sockerrno()
        GoTo error_
    End If

    ws_freeaddrinfo ai_list
    udp_create_socket = sock
    Exit Function

error_:
    ws_freeaddrinfo ai_list
    If sock <> INVALID_SOCKET Then closesocket sock

    udp_create_socket = INVALID_SOCKET
End Function

' int udp_recvfrom(socket_t sock, char *buffer, size_t size, addr_record_t *src)
Public Function udp_recvfrom(ByVal sock As Long, buffer() As Byte, ByVal size As Long, src As addr_record_t) As Long
    Dim len_ As Long
    src.len_ = 128                              ' sizeof(src->addr)
    len_ = ws_recvfrom(sock, buffer(0), size, 0, src.addr(0), src.len_)
    plum_sock_called
    If len_ >= 0 Then addr_unmap_inet6_v4mapped src.addr, src.len_

    udp_recvfrom = len_
End Function

' int udp_sendto(socket_t sock, const char *data, size_t size, const addr_record_t *dst)
Public Function udp_sendto(ByVal sock As Long, data() As Byte, ByVal size As Long, dst As addr_record_t) As Long
    Dim tmp As addr_record_t, name_ As addr_record_t
    tmp = dst
    name_.len_ = 128
    If getsockname(sock, name_.addr, name_.len_) = 0 Then
        If sa_family(name_.addr) = AF_INET6 Then addr_map_inet6_v4mapped tmp.addr, tmp.len_
    Else
        plum_log_write PLUM_LOG_LEVEL_WARN, "getsockname failed, errno=%d", sockerrno()
    End If
    udp_sendto = sendto_(sock, data, size, tmp.addr, tmp.len_)
End Function

' int udp_sendto_self(socket_t sock, const char *data, size_t size)
Public Function udp_sendto_self(ByVal sock As Long, data() As Byte, ByVal size As Long) As Long
    Dim local_ As addr_record_t, ret As Long
    If udp_get_local_addr(sock, AF_UNSPEC, local_) < 0 Then
        udp_sendto_self = -1
        Exit Function
    End If

    ' We know local has the same address family as sock here
    ret = sendto_(sock, data, size, local_.addr, local_.len_)
    If ret >= 0 Or sa_family(local_.addr) <> AF_INET6 Then
        udp_sendto_self = ret
        Exit Function
    End If

    ' Fallback as IPv6 may be disabled on the loopback interface
    If udp_get_local_addr(sock, AF_INET, local_) < 0 Then
        udp_sendto_self = -1
        Exit Function
    End If

    addr_map_inet6_v4mapped local_.addr, local_.len_
    udp_sendto_self = sendto_(sock, data, size, local_.addr, local_.len_)
End Function

' uint16_t udp_get_port(socket_t sock)
Public Function udp_get_port(ByVal sock As Long) As Long
    Dim record As addr_record_t
    If udp_get_bound_addr(sock, record) < 0 Then
        udp_get_port = 0
        Exit Function
    End If
    udp_get_port = addr_get_port(record.addr)
End Function

' int udp_get_bound_addr(socket_t sock, addr_record_t *record)
Public Function udp_get_bound_addr(ByVal sock As Long, record As addr_record_t) As Long
    record.len_ = 128
    If getsockname(sock, record.addr, record.len_) <> 0 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "getsockname failed, errno=%d", sockerrno()
        udp_get_bound_addr = -1
        Exit Function
    End If
    udp_get_bound_addr = 0
End Function

' int udp_get_local_addr(socket_t sock, int family_hint, addr_record_t *record): family_hint may be AF_UNSPEC
Public Function udp_get_local_addr(ByVal sock As Long, ByVal family_hint As Long, record As addr_record_t) As Long
    Dim port As Long, i As Long
    If udp_get_bound_addr(sock, record) < 0 Then
        udp_get_local_addr = -1
        Exit Function
    End If

    ' If the socket is bound to a particular address, return it
    If Not addr_is_any(record.addr) Then
        If sa_family(record.addr) = AF_INET And family_hint = AF_INET6 Then addr_map_inet6_v4mapped record.addr, record.len_

        udp_get_local_addr = 0
        Exit Function
    End If

    If sa_family(record.addr) = AF_INET6 And family_hint = AF_INET Then
        ' Generate an IPv4 instead (socket is listening to any IPv4 or IPv6)

        port = addr_get_port(record.addr)
        If port = 0 Then
            udp_get_local_addr = -1
            Exit Function
        End If

        For i = 0 To 15                         ' memset(sin, 0, sizeof(*sin));
            record.addr(i) = 0
        Next
        record.addr(0) = AF_INET
        addr_set_port record.addr, port
        record.len_ = 16
    End If

    Select Case sa_family(record.addr)
    Case AF_INET
        record.addr(4) = 127: record.addr(5) = 0: record.addr(6) = 0: record.addr(7) = 1   ' localhost
    Case AF_INET6
        For i = 0 To 14
            record.addr(8 + i) = 0
        Next
        record.addr(8 + 15) = 1                 ' localhost
    Case Else
        ' Ignore
    End Select

    If sa_family(record.addr) = AF_INET And family_hint = AF_INET6 Then addr_map_inet6_v4mapped record.addr, record.len_

    udp_get_local_addr = 0
End Function
