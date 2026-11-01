Attribute VB_Name = "plum_addr_c"
' libplum: Copyright (c) 2020-2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/addr.c, addr.h. A struct sockaddr (sockaddr_storage, sockaddr_in, sockaddr_in6) is its bytes: a
' Byte array (an addr_record_t's addr); a socklen_t * is a ByRef Long. A char buffer of size bytes is a ByRef String
' that gets at most size - 1 characters (snprintf's); the functions return what C's return. addr_is_temp_inet6
' (used by no caller) is left out.
Option Explicit

Private Declare Function ws_inet_ntop Lib "ws2_32" Alias "inet_ntop" (ByVal Family As Long, pAddr As Any, ByVal pStringBuf As String, ByVal StringBufSize As Long) As Long
Private Declare Function ws_inet_pton Lib "ws2_32" Alias "inet_pton" (ByVal Family As Long, ByVal pszAddrString As String, pAddrBuf As Any) As Long
Private Declare Function ws_getnameinfo Lib "ws2_32" Alias "getnameinfo" (pSockaddr As Any, ByVal SockaddrLength As Long, ByVal pNodeBuffer As String, ByVal NodeBufferSize As Long, ByVal pServiceBuffer As String, ByVal ServiceBufferSize As Long, ByVal Flags As Long) As Long
Private Declare Function ws_getaddrinfo Lib "ws2_32" Alias "getaddrinfo" (ByVal pNodeName As String, ByVal pServiceName As String, pHints As ADDRINFOA, ppResult As Long) As Long
Private Declare Sub ws_freeaddrinfo Lib "ws2_32" Alias "freeaddrinfo" (ByVal pAddrInfo As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Const AF_UNSPEC As Long = 0
Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23
Private Const SOCK_DGRAM As Long = 2
Private Const IPPROTO_UDP As Long = 17
Private Const AI_NUMERICHOST As Long = &H4
Private Const AI_NUMERICSERV As Long = &H8
Private Const AI_ADDRCONFIG As Long = &H400
Private Const NI_NUMERICHOST As Long = &H2
Private Const NI_NUMERICSERV As Long = &H8
Private Const NI_DGRAM As Long = &H10
Private Const INET_ADDRSTRLEN As Long = 22
Private Const INET6_ADDRSTRLEN As Long = 65
Private Const SOCKADDR_IN_SIZE As Long = 16     ' sizeof(struct sockaddr_in)
Private Const SOCKADDR_IN6_SIZE As Long = 28    ' sizeof(struct sockaddr_in6)

' IPv6 max representation length is 45 plus 4 for potential zone index
Public Const ADDR_MAX_NUMERICHOST_LEN As Long = 56  ' 45 + 4 + 1 rounded up
Public Const ADDR_MAX_NUMERICSERV_LEN As Long = 8   ' 5 + 1 rounded up
Public Const ADDR_MAX_STRING_LEN As Long = 64

' typedef struct addr_record
Public Type addr_record_t
    addr(0 To 127) As Byte                      ' struct sockaddr_storage addr: sa_family at 0, the port (network
                                                ' order) at 2, sin_addr at 4, sin6_addr at 8
    len_ As Long                                ' socklen_t len
End Type

' (VB) sa->sa_family
Public Function sa_family(sa() As Byte) As Long
    sa_family = CLng(sa(0)) + CLng(sa(1)) * 256
End Function

' (VB) a snprintf(buffer, size, "%s", s): the length of s
Private Function snprintf_s(buffer As String, ByVal size As Long, ByVal s As String) As Long
    If Len(s) > size - 1 Then buffer = Left$(s, size - 1) Else buffer = s
    snprintf_s = Len(s)
End Function

' (VB) a char[] Winsock filled: the text before its NUL
Private Function cstr_(ByVal s As String) As String
    Dim i As Long
    i = InStr(s, vbNullChar)
    If i > 0 Then cstr_ = Left$(s, i - 1) Else cstr_ = s
End Function

' (VB) IN6_IS_ADDR_V4MAPPED(&sin6->sin6_addr)
Private Function in6_is_v4mapped(sa() As Byte) As Boolean
    Dim i As Long
    For i = 8 To 17
        If sa(i) <> 0 Then Exit Function
    Next
    in6_is_v4mapped = sa(18) = &HFF And sa(19) = &HFF
End Function

' socklen_t addr_get_len(const struct sockaddr *sa)
Public Function addr_get_len(sa() As Byte) As Long
    Select Case sa_family(sa)
    Case AF_INET
        addr_get_len = SOCKADDR_IN_SIZE
    Case AF_INET6
        addr_get_len = SOCKADDR_IN6_SIZE
    Case Else
        plum_log_write PLUM_LOG_LEVEL_WARN, "Unknown address family %hu", sa_family(sa)
        addr_get_len = 0
    End Select
End Function

' int addr_get_host(const struct sockaddr *sa, char *buffer, size_t size)
Public Function addr_get_host(sa() As Byte, buffer As String, ByVal size As Long) As Long
    Dim tmp As String
    Select Case sa_family(sa)
    Case AF_INET
        tmp = String$(INET_ADDRSTRLEN, vbNullChar)
        If inet_ntop(AF_INET, sa, 4, tmp, INET_ADDRSTRLEN) = 0 Then
            plum_log_write PLUM_LOG_LEVEL_ERROR, "inet_ntop failed, errno=%d", sockerrno()
            GoTo error_
        End If
        addr_get_host = snprintf_s(buffer, size, cstr_(tmp))
        Exit Function
    Case AF_INET6
        tmp = String$(INET6_ADDRSTRLEN, vbNullChar)
        If inet_ntop(AF_INET6, sa, 8, tmp, INET6_ADDRSTRLEN) = 0 Then
            plum_log_write PLUM_LOG_LEVEL_ERROR, "inet_ntop failed, errno=%d", sockerrno()
            GoTo error_
        End If
        addr_get_host = snprintf_s(buffer, size, cstr_(tmp))
        Exit Function
    Case Else
        plum_log_write PLUM_LOG_LEVEL_WARN, "Unknown address family %hu", sa_family(sa)
        GoTo error_
    End Select

error_:
    ' Make sure we still write a valid null-terminated string
    snprintf_s buffer, size, "?"
    addr_get_host = -1
End Function

' (VB) inet_ntop(family, sa + offset, buf, size), through plum_sock_called
Private Function inet_ntop(ByVal family As Long, sa() As Byte, ByVal offset As Long, buf As String, ByVal size As Long) As Long
    inet_ntop = ws_inet_ntop(family, sa(offset), buf, size)
    plum_sock_called
End Function

' (VB) inet_pton(family, host, sa + offset)
Private Function inet_pton(ByVal family As Long, ByVal host As String, sa() As Byte, ByVal offset As Long) As Long
    inet_pton = ws_inet_pton(family, host, sa(offset))
    plum_sock_called
End Function

' uint16_t addr_get_port(const struct sockaddr *sa)
Public Function addr_get_port(sa() As Byte) As Long
    Select Case sa_family(sa)
    Case AF_INET
        addr_get_port = get_be16(sa, 2)         ' ntohs(((struct sockaddr_in *)sa)->sin_port)
    Case AF_INET6
        addr_get_port = get_be16(sa, 2)         ' ntohs(((struct sockaddr_in6 *)sa)->sin6_port)
    Case Else
        plum_log_write PLUM_LOG_LEVEL_WARN, "Unknown address family %hu", sa_family(sa)
        addr_get_port = 0
    End Select
End Function

' int addr_set_port(struct sockaddr *sa, uint16_t port)
Public Function addr_set_port(sa() As Byte, ByVal port As Long) As Long
    Select Case sa_family(sa)
    Case AF_INET
        put_be16 sa, 2, port
        addr_set_port = 0
    Case AF_INET6
        put_be16 sa, 2, port
        addr_set_port = 0
    Case Else
        plum_log_write PLUM_LOG_LEVEL_WARN, "Unknown address family %hu", sa_family(sa)
        addr_set_port = -1
    End Select
End Function

' bool addr_is_any(const struct sockaddr *sa)
Public Function addr_is_any(sa() As Byte) As Boolean
    Dim i As Long
    Select Case sa_family(sa)
    Case AF_INET
        For i = 0 To 3
            If sa(4 + i) <> 0 Then Exit Function
        Next
        addr_is_any = True
    Case AF_INET6
        If in6_is_v4mapped(sa) Then
            For i = 0 To 3
                If sa(8 + 12 + i) <> 0 Then Exit Function
            Next
        Else
            For i = 0 To 15
                If sa(8 + i) <> 0 Then Exit Function
            Next
        End If
        addr_is_any = True
    Case Else
        addr_is_any = False
    End Select
End Function

' bool addr_is_local(const struct sockaddr *sa)
Public Function addr_is_local(sa() As Byte) As Boolean
    Dim i As Long
    Select Case sa_family(sa)
    Case AF_INET
        If sa(4) = 127 Then addr_is_local = True: Exit Function                     ' loopback
        If sa(4) = 169 And sa(5) = 254 Then addr_is_local = True: Exit Function     ' link-local
        addr_is_local = False
    Case AF_INET6
        ' IN6_IS_ADDR_LOOPBACK(&sin6->sin6_addr)
        For i = 0 To 14
            If sa(8 + i) <> 0 Then Exit For
        Next
        If i = 15 And sa(8 + 15) = 1 Then addr_is_local = True: Exit Function
        ' IN6_IS_ADDR_LINKLOCAL(&sin6->sin6_addr)
        If sa(8) = &HFE And (sa(9) And &HC0) = &H80 Then addr_is_local = True: Exit Function
        If in6_is_v4mapped(sa) Then
            If sa(8 + 12) = 127 Then addr_is_local = True: Exit Function                    ' loopback
            If sa(8 + 12) = 169 And sa(8 + 13) = 254 Then addr_is_local = True: Exit Function ' link-local
            addr_is_local = False
            Exit Function
        End If
        addr_is_local = False
    Case Else
        addr_is_local = False
    End Select
End Function

' bool addr_is_private(const struct sockaddr *sa)
Public Function addr_is_private(sa() As Byte) As Boolean
    Select Case sa_family(sa)
    Case AF_INET
        ' RFC 1918 Private Addresses
        If sa(4) = 10 Then addr_is_private = True: Exit Function                            ' 10.0.0.0/8
        If sa(4) = 172 And (sa(5) And &HF0) = 16 Then addr_is_private = True: Exit Function ' 172.16.0.0/12
        If sa(4) = 192 And sa(5) = 168 Then addr_is_private = True: Exit Function           ' 192.168.0.0/16
        addr_is_private = False
    Case AF_INET6
        ' RFC 4193 Unique Local Addresses
        If (sa(8) And &HFE) = &HFC Then addr_is_private = True: Exit Function              ' fc00::/7
        addr_is_private = False
    Case Else
        addr_is_private = False
    End Select
End Function

' bool addr_is_public(const struct sockaddr *sa)
Public Function addr_is_public(sa() As Byte) As Boolean
    addr_is_public = Not addr_is_private(sa) And Not addr_is_local(sa)
End Function

' bool addr_unmap_inet6_v4mapped(struct sockaddr *sa, socklen_t *len)
Public Function addr_unmap_inet6_v4mapped(sa() As Byte, len_ As Long) As Boolean
    Dim copy(0 To 27) As Byte, i As Long
    If sa_family(sa) <> AF_INET6 Then Exit Function

    If Not in6_is_v4mapped(sa) Then Exit Function

    For i = 0 To 27                             ' struct sockaddr_in6 copy = *sin6;
        copy(i) = sa(i)
    Next

    For i = 0 To SOCKADDR_IN_SIZE - 1           ' memset(sin, 0, sizeof(*sin));
        sa(i) = 0
    Next
    sa(0) = AF_INET                             ' sin->sin_family = AF_INET;
    sa(2) = copy(2): sa(3) = copy(3)            ' sin->sin_port = sin6->sin6_port;
    For i = 0 To 3                              ' memcpy(&sin->sin_addr, ((const uint8_t *)&sin6->sin6_addr) + 12, 4);
        sa(4 + i) = copy(8 + 12 + i)
    Next
    len_ = SOCKADDR_IN_SIZE
    addr_unmap_inet6_v4mapped = True
End Function

' bool addr_map_inet6_v4mapped(struct sockaddr_storage *ss, socklen_t *len)
Public Function addr_map_inet6_v4mapped(ss() As Byte, len_ As Long) As Boolean
    Dim copy(0 To 15) As Byte, i As Long
    If sa_family(ss) <> AF_INET Then Exit Function

    For i = 0 To 15                             ' struct sockaddr_in copy = *sin;
        copy(i) = ss(i)
    Next

    For i = 0 To SOCKADDR_IN6_SIZE - 1          ' memset(sin6, 0, sizeof(*sin6));
        ss(i) = 0
    Next
    ss(0) = AF_INET6 And &HFF&                  ' sin6->sin6_family = AF_INET6;
    ss(1) = AF_INET6 \ 256
    ss(2) = copy(2): ss(3) = copy(3)            ' sin6->sin6_port = sin->sin_port;
    ' memset(b, 0, 10); memset(b + 10, 0xFF, 2); memcpy(b + 12, (const uint8_t *)&sin->sin_addr, 4);
    ss(8 + 10) = &HFF
    ss(8 + 11) = &HFF
    For i = 0 To 3
        ss(8 + 12 + i) = copy(4 + i)
    Next
    len_ = SOCKADDR_IN6_SIZE
    addr_map_inet6_v4mapped = True
End Function

' bool addr_is_equal(const struct sockaddr *a, const struct sockaddr *b, bool compare_ports)
Public Function addr_is_equal(a() As Byte, b() As Byte, ByVal compare_ports As Boolean) As Boolean
    Dim i As Long
    If sa_family(a) <> sa_family(b) Then Exit Function

    Select Case sa_family(a)
    Case AF_INET
        For i = 4 To 7
            If a(i) <> b(i) Then Exit Function
        Next
        If compare_ports And (a(2) <> b(2) Or a(3) <> b(3)) Then Exit Function
    Case AF_INET6
        For i = 8 To 23
            If a(i) <> b(i) Then Exit Function
        Next
        If compare_ports And (a(2) <> b(2) Or a(3) <> b(3)) Then Exit Function
    Case Else
        Exit Function
    End Select

    addr_is_equal = True
End Function

' int addr_to_string(const struct sockaddr *sa, char *buffer, size_t size)
Public Function addr_to_string(sa() As Byte, buffer As String, ByVal size As Long) As Long
    Dim salen As Long, host As String, service As String, len_ As Long, s As String
    salen = addr_get_len(sa)
    If salen = 0 Then GoTo error_

    host = String$(ADDR_MAX_NUMERICHOST_LEN, vbNullChar)
    service = String$(ADDR_MAX_NUMERICSERV_LEN, vbNullChar)
    If getnameinfo(sa, salen, host, ADDR_MAX_NUMERICHOST_LEN, service, ADDR_MAX_NUMERICSERV_LEN, _
            NI_NUMERICHOST Or NI_NUMERICSERV Or NI_DGRAM) <> 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "getnameinfo failed, errno=%d", sockerrno()
        GoTo error_
    End If
    host = cstr_(host)
    service = cstr_(service)

    If service = "0" Then
        s = host
    Else
        s = host & ":" & service
    End If
    len_ = Len(s)

    If len_ < 0 Or len_ >= size Then GoTo error_
    buffer = s

    addr_to_string = len_
    Exit Function

error_:
    ' Make sure we still write a valid null-terminated string
    snprintf_s buffer, size, "?"
    addr_to_string = -1
End Function

Private Function getnameinfo(sa() As Byte, ByVal salen As Long, host As String, ByVal hostlen As Long, service As String, ByVal servlen As Long, ByVal flags As Long) As Long
    getnameinfo = ws_getnameinfo(sa(0), salen, host, hostlen, service, servlen, flags)
    plum_sock_called
End Function

Private Function getaddrinfo(ByVal node As String, ByVal service As String, hints As ADDRINFOA, res As Long) As Long
    getaddrinfo = ws_getaddrinfo(node, service, hints, res)
    plum_sock_called
End Function

' int addr_set(int family, const char *host, uint16_t port, addr_record_t *record)
Public Function addr_set(ByVal family As Long, ByVal host As String, ByVal port As Long, record As addr_record_t) As Long
    Dim i As Long, hints As ADDRINFOA, zeroHints As ADDRINFOA, service As String, ai_list As Long, ai As Long, info As ADDRINFOA
    Select Case family
    Case AF_INET
        For i = 0 To SOCKADDR_IN_SIZE - 1           ' memset(sin, 0, sizeof(*sin));
            record.addr(i) = 0
        Next
        record.len_ = SOCKADDR_IN_SIZE
        record.addr(0) = AF_INET
        put_be16 record.addr, 2, port
        If inet_pton(AF_INET, host, record.addr, 4) <> 1 Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Invalid IPv4 address: %s", host
            addr_set = -1
            Exit Function
        End If
    Case AF_INET6
        For i = 0 To SOCKADDR_IN6_SIZE - 1          ' memset(sin6, 0, sizeof(*sin6));
            record.addr(i) = 0
        Next
        record.len_ = SOCKADDR_IN6_SIZE
        record.addr(0) = AF_INET6 And &HFF&
        record.addr(1) = AF_INET6 \ 256
        put_be16 record.addr, 2, port
        If inet_pton(AF_INET6, host, record.addr, 8) <> 1 Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Invalid IPv6 address: %s", host
            addr_set = -1
            Exit Function
        End If
    Case AF_UNSPEC
        hints = zeroHints
        hints.ai_family = AF_UNSPEC
        hints.ai_socktype = SOCK_DGRAM
        hints.ai_protocol = IPPROTO_UDP
        hints.ai_flags = AI_NUMERICHOST Or AI_NUMERICSERV
        service = CStr(port)

        ai_list = 0
        If getaddrinfo(host, service, hints, ai_list) <> 0 Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Invalid address: %s", host
            addr_set = -1
            Exit Function
        End If

        ai = ai_list
        Do While ai <> 0
            CopyMemory info, ByVal ai, LenB(info)
            If info.ai_family = AF_INET Or info.ai_family = AF_INET6 Then Exit Do
            ai = info.ai_next
        Loop

        If ai = 0 Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Not an IPv4 or IPv6 address: %s", host
            ws_freeaddrinfo ai_list
            addr_set = -1
            Exit Function
        End If

        CopyMemory record.addr(0), ByVal info.ai_addr, info.ai_addrlen
        record.len_ = info.ai_addrlen
        ws_freeaddrinfo ai_list
    Case Else
        plum_log_write PLUM_LOG_LEVEL_WARN, "Unknown address family %d", family
        addr_set = -1
        Exit Function
    End Select

    addr_set = 0
End Function

' int addr_set_binary(int family, const void *addr, uint16_t port, addr_record_t *record): addr is the bytes at
' offset in addr()
Public Function addr_set_binary(ByVal family As Long, addr() As Byte, ByVal offset As Long, ByVal port As Long, record As addr_record_t) As Long
    Dim i As Long
    Select Case family
    Case AF_INET
        For i = 0 To SOCKADDR_IN_SIZE - 1
            record.addr(i) = 0
        Next
        record.len_ = SOCKADDR_IN_SIZE
        record.addr(0) = AF_INET
        put_be16 record.addr, 2, port
        For i = 0 To 3
            record.addr(4 + i) = addr(offset + i)
        Next
    Case AF_INET6
        For i = 0 To SOCKADDR_IN6_SIZE - 1
            record.addr(i) = 0
        Next
        record.len_ = SOCKADDR_IN6_SIZE
        record.addr(0) = AF_INET6 And &HFF&
        record.addr(1) = AF_INET6 \ 256
        put_be16 record.addr, 2, port
        For i = 0 To 15
            record.addr(8 + i) = addr(offset + i)
        Next
    Case AF_UNSPEC
        plum_log_write PLUM_LOG_LEVEL_ERROR, "addr_set_binary called with unspecified address family"
        addr_set_binary = -1
        Exit Function
    Case Else
        plum_log_write PLUM_LOG_LEVEL_WARN, "Unknown address family %d", family
        addr_set_binary = -1
        Exit Function
    End Select

    addr_set_binary = 0
End Function

' int addr_resolve(const char *hostname, const char *service, addr_record_t *records, size_t count)
Public Function addr_resolve(ByVal hostname As String, ByVal service As String, records() As addr_record_t, ByVal count As Long) As Long
    Dim hints As ADDRINFOA, ai_list As Long, ai As Long, info As ADDRINFOA, ret As Long, r As Long
    r = 0                                       ' addr_record_t *end = records + count;

    hints.ai_family = AF_UNSPEC
    hints.ai_socktype = SOCK_DGRAM
    hints.ai_protocol = IPPROTO_UDP
    hints.ai_flags = AI_ADDRCONFIG
    ai_list = 0
    If getaddrinfo(hostname, service, hints, ai_list) <> 0 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Address resolution failed for %s:%s", hostname, service
        addr_resolve = -1
        Exit Function
    End If

    ret = 0
    ai = ai_list
    Do While ai <> 0
        CopyMemory info, ByVal ai, LenB(info)
        If info.ai_family = AF_INET Or info.ai_family = AF_INET6 Then
            ret = ret + 1
            If r <> count Then
                CopyMemory records(r).addr(0), ByVal info.ai_addr, info.ai_addrlen
                records(r).len_ = info.ai_addrlen
                r = r + 1
            End If
        End If
        ai = info.ai_next
    Loop

    ws_freeaddrinfo ai_list
    addr_resolve = ret
End Function

' bool addr_record_is_equal(const addr_record_t *a, const addr_record_t *b, bool compare_ports)
Public Function addr_record_is_equal(a As addr_record_t, b As addr_record_t, ByVal compare_ports As Boolean) As Boolean
    addr_record_is_equal = addr_is_equal(a.addr, b.addr, compare_ports)
End Function

' int addr_record_to_string(const addr_record_t *record, char *buffer, size_t size)
Public Function addr_record_to_string(record As addr_record_t, buffer As String, ByVal size As Long) As Long
    addr_record_to_string = addr_to_string(record.addr, buffer, size)
End Function
