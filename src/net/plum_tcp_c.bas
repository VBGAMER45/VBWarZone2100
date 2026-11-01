Attribute VB_Name = "plum_tcp_c"
' libplum: Copyright (c) 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/tcp.c, tcp.h (the _WIN32 branches). A buffer is a Byte array and the offset of its first byte.
' (VB) The three waits are resumable (plum_socket_h's PLUM_PENDING): a poll() with the time left as its timeout is
' a poll with timeout 0 that returns PLUM_PENDING while time is left; the call that got PLUM_PENDING is made again on
' the next pass with the same arguments, and goes on where it stopped (tcp_connect_socket's and tcp_send's state is
' kept here; tcp_recv has none to keep).
Option Explicit

Private Declare Function ws_socket Lib "ws2_32" Alias "socket" (ByVal af As Long, ByVal type_ As Long, ByVal protocol As Long) As Long
Private Declare Function ws_closesocket Lib "ws2_32" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function ws_connect Lib "ws2_32" Alias "connect" (ByVal s As Long, name_ As Any, ByVal namelen As Long) As Long
Private Declare Function ws_ioctlsocket Lib "ws2_32" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function ws_getsockopt Lib "ws2_32" Alias "getsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function ws_recv Lib "ws2_32" Alias "recv" (ByVal s As Long, buf As Any, ByVal len_ As Long, ByVal flags As Long) As Long
Private Declare Function ws_send Lib "ws2_32" Alias "send" (ByVal s As Long, buf As Any, ByVal len_ As Long, ByVal flags As Long) As Long

Private Const INVALID_SOCKET As Long = -1
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const SOL_SOCKET As Long = &HFFFF&
Private Const SO_ERROR As Long = &H1007
Private Const FIONBIO As Long = &H8004667E

Public Const TCP_ERR_UNKNOWN As Long = -1
Public Const TCP_ERR_TIMEOUT As Long = -2

' (VB) the resumable calls' state (one call of each at a time: the client thread's)
Private connect_at As Long                      ' tcp_connect_socket: 1 = waiting for the connection
Private connect_sock As Long
Private send_at As Long                         ' tcp_send: 1 = in its loop
Private send_left As Long
Private send_data As Long                       ' the offset of the data left

Private Function closesocket(ByVal s As Long) As Long
    closesocket = ws_closesocket(s)
    plum_sock_called
End Function

' socket_t tcp_connect_socket(const addr_record_t *remote_addr, timestamp_t end_timestamp)
Public Function tcp_connect_socket(remote_addr As addr_record_t, ByVal end_timestamp As Double) As Long
    Dim family As Long, nbio As Long, ret As Long, pfd(0 To 0) As pollfd, timediff As Double, err_ As Long, errlen As Long
    If connect_at = 1 Then connect_at = 0: GoTo wait_  ' (VB) resumed

    connect_sock = INVALID_SOCKET
    family = sa_family(remote_addr.addr)

    ' Create socket
    connect_sock = ws_socket(family, SOCK_STREAM, IPPROTO_TCP)
    plum_sock_called
    If connect_sock = INVALID_SOCKET Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "TCP socket creation failed, errno=%d", sockerrno()
        GoTo error_
    End If

    nbio = 1
    If ws_ioctlsocket(connect_sock, FIONBIO, nbio) <> 0 Then
        plum_sock_called
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Setting non-blocking mode on TCP socket failed, errno=%d", sockerrno()
        GoTo error_
    End If

    ' Initiate connection
    ret = ws_connect(connect_sock, remote_addr.addr(0), remote_addr.len_)
    plum_sock_called
    If ret <> 0 And sockerrno() <> SEINPROGRESS And sockerrno() <> SEWOULDBLOCK Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "TCP connection failed, errno=%d", sockerrno()
        GoTo error_
    End If

    ' Wait for connection
wait_:
    pfd(0).fd = connect_sock
    pfd(0).events = POLLOUT

    Do
        timediff = end_timestamp - current_timestamp()
        If timediff < 0 Then timediff = 0

        ret = poll(pfd, 1, 0)                   ' poll(&pfd, 1, (int)timediff)
        If ret = 0 And timediff > 0 Then connect_at = 1: tcp_connect_socket = PLUM_PENDING: Exit Function   ' (VB)
    Loop While ret < 0 And (sockerrno() = SEINTR Or sockerrno() = SEAGAIN)

    If ret < 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Failed to wait for TCP socket connection"
        GoTo error_
    End If

    If (pfd(0).revents And POLLOUT) = 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "TCP connection timed out"
        GoTo error_
    End If

    err_ = 0
    errlen = 4
    If ws_getsockopt(connect_sock, SOL_SOCKET, SO_ERROR, err_, errlen) <> 0 Then
        plum_sock_called
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Failed to get socket error code"
        GoTo error_
    End If

    If err_ <> 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "TCP connection failed, errno=%d", err_
        GoTo error_
    End If

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "TCP connection succeeded"
    tcp_connect_socket = connect_sock
    Exit Function

error_:
    If connect_sock <> INVALID_SOCKET Then closesocket connect_sock

    tcp_connect_socket = INVALID_SOCKET
End Function

' int tcp_recv(socket_t sock, char *buffer, size_t size, timestamp_t end_timestamp): into buffer from offset
Public Function tcp_recv(ByVal sock As Long, buffer() As Byte, ByVal offset As Long, ByVal size As Long, ByVal end_timestamp As Double) As Long
    Dim timediff As Double, pfd(0 To 0) As pollfd, ret As Long, flags As Long, len_ As Long
    Do While True
        timediff = end_timestamp - current_timestamp()
        If timediff < 0 Then timediff = 0

        pfd(0).fd = sock
        pfd(0).events = POLLIN

        ret = poll(pfd, 1, 0)                   ' poll(&pfd, 1, (int)timediff)
        If ret = 0 And timediff > 0 Then tcp_recv = PLUM_PENDING: Exit Function   ' (VB)
        If ret < 0 Then
            If sockerrno() = SEINTR Or sockerrno() = SEAGAIN Then
                plum_log_write PLUM_LOG_LEVEL_VERBOSE, "poll interrupted"
                GoTo continue_
            Else
                plum_log_write PLUM_LOG_LEVEL_ERROR, "poll failed, errno=%d", sockerrno()
                tcp_recv = TCP_ERR_UNKNOWN
                Exit Function
            End If
        End If

        If ret = 0 Then Exit Do ' timeout

        If (pfd(0).revents And POLLNVAL) <> 0 Or (pfd(0).revents And POLLERR) <> 0 Then
            plum_log_write PLUM_LOG_LEVEL_ERROR, "Error when polling socket"
            tcp_recv = TCP_ERR_UNKNOWN
            Exit Function
        End If

        If (pfd(0).revents And POLLIN) <> 0 Or (pfd(0).revents And POLLHUP) <> 0 Then
            flags = 0
            len_ = ws_recv(pfd(0).fd, buffer(offset), size, flags)
            plum_sock_called
            If len_ < 0 Then
                If sockerrno() = SEAGAIN Or sockerrno() = SEWOULDBLOCK Then GoTo continue_

                plum_log_write PLUM_LOG_LEVEL_WARN, "TCP recv failed, errno=%d", sockerrno()
                tcp_recv = TCP_ERR_UNKNOWN
                Exit Function
            End If

            tcp_recv = len_
            Exit Function
        End If
continue_:
    Loop

    ' Timeout
    plum_log_write PLUM_LOG_LEVEL_WARN, "TCP recv timeout"
    tcp_recv = TCP_ERR_TIMEOUT
End Function

' int tcp_send(socket_t sock, const char *data, size_t size, timestamp_t end_timestamp): data from offset
Public Function tcp_send(ByVal sock As Long, data() As Byte, ByVal offset As Long, ByVal size As Long, ByVal end_timestamp As Double) As Long
    Dim timediff As Double, pfd(0 To 0) As pollfd, ret As Long, flags As Long, len_ As Long, sent As Long
    If send_at = 1 Then
        send_at = 0                             ' (VB) resumed: send_left and send_data as they were
    Else
        send_left = size
        send_data = offset
    End If
    Do While True
        timediff = end_timestamp - current_timestamp()
        If timediff < 0 Then timediff = 0

        pfd(0).fd = sock
        pfd(0).events = POLLOUT

        ret = poll(pfd, 1, 0)                   ' poll(&pfd, 1, (int)timediff)
        If ret = 0 And timediff > 0 Then send_at = 1: tcp_send = PLUM_PENDING: Exit Function   ' (VB)
        If ret < 0 Then
            If sockerrno() = SEINTR Or sockerrno() = SEAGAIN Then
                plum_log_write PLUM_LOG_LEVEL_VERBOSE, "poll interrupted"
                GoTo continue_
            Else
                plum_log_write PLUM_LOG_LEVEL_ERROR, "poll failed, errno=%d", sockerrno()
                tcp_send = TCP_ERR_UNKNOWN
                Exit Function
            End If
        End If

        If ret = 0 Then Exit Do

        If (pfd(0).revents And POLLNVAL) <> 0 Or (pfd(0).revents And POLLERR) <> 0 Then
            plum_log_write PLUM_LOG_LEVEL_ERROR, "Error when polling socket"
            tcp_send = TCP_ERR_UNKNOWN
            Exit Function
        End If

        If (pfd(0).revents And POLLOUT) <> 0 Then
            flags = 0
            len_ = ws_send(pfd(0).fd, data(send_data), size, flags)    ' send(pfd.fd, data, (int)size, flags)
            plum_sock_called
            If len_ < 0 Then
                If sockerrno() = SEAGAIN Or sockerrno() = SEWOULDBLOCK Then GoTo continue_

                plum_log_write PLUM_LOG_LEVEL_WARN, "TCP send failed, errno=%d", sockerrno()
                tcp_send = TCP_ERR_UNKNOWN
                Exit Function
            End If

            send_data = send_data + len_
            send_left = send_left - len_

            If send_left = 0 Then
                tcp_send = size
                Exit Function
            End If
        End If
continue_:
    Loop

    ' Timeout
    plum_log_write PLUM_LOG_LEVEL_WARN, "TCP send timeout"
    sent = size - send_left
    If sent = 0 Then tcp_send = TCP_ERR_TIMEOUT Else tcp_send = sent
End Function
