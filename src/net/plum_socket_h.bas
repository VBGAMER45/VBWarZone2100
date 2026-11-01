Attribute VB_Name = "plum_socket_h"
' libplum: Copyright (c) 2020 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum (ref/libplum, commit 47be21b as 4.7.0 pins it; MPL 2.0) src/socket.h, its _WIN32 branch: Winsock 2.
' poll is WSAPoll. sockerrno (WSAGetLastError()) is the last Winsock error as the libplum modules' wrappers keep it:
' VB clears the thread's last error before each declared call, so each wrapper calls plum_sock_called right after
' its call (a call that succeeds leaves the older code, as WSAGetLastError does).
' (VB) No threads (docs/CONVENTIONS.md): libplum's client thread is a poll (plum_c's plum_poll). A call that would
' wait (poll() with a timeout, a condition variable) returns PLUM_PENDING instead and keeps its state; its caller
' keeps its own and returns PLUM_PENDING too, up to the thread's pass, which the next pass resumes (plum_client_c).
Option Explicit

Private Declare Function ws_WSAPoll Lib "ws2_32" Alias "WSAPoll" (fdArray As Any, ByVal fds As Long, ByVal timeout As Long) As Long

' #define SE... WSAE...
Public Const SEADDRINUSE As Long = 10048        ' WSAEADDRINUSE
Public Const SEINTR As Long = 10004             ' WSAEINTR
Public Const SEAGAIN As Long = 10035            ' WSAEWOULDBLOCK
Public Const SEACCES As Long = 10013            ' WSAEACCES
Public Const SEWOULDBLOCK As Long = 10035       ' WSAEWOULDBLOCK
Public Const SEINPROGRESS As Long = 10036       ' WSAEINPROGRESS
Public Const SECONNREFUSED As Long = 10061      ' WSAECONNREFUSED
Public Const SECONNRESET As Long = 10054        ' WSAECONNRESET
Public Const SENETRESET As Long = 10052         ' WSAENETRESET

' winsock2.h: WSAPOLLFD (struct pollfd) and its events
Public Type pollfd
    fd As Long
    events As Integer
    revents As Integer
End Type
Public Const POLLERR As Long = &H1
Public Const POLLHUP As Long = &H2
Public Const POLLNVAL As Long = &H4
Public Const POLLWRNORM As Long = &H10
Public Const POLLOUT As Long = &H10             ' POLLWRNORM
Public Const POLLRDNORM As Long = &H100
Public Const POLLRDBAND As Long = &H200
Public Const POLLIN As Long = &H300             ' POLLRDNORM | POLLRDBAND

' (VB) the call would wait: the client thread's pass ends here and the next one resumes the call (see above)
Public Const PLUM_PENDING As Long = -1000

Private sockerrno_ As Long

' (VB) keeps the error code the Winsock call just made set (Err.LastDllError), as WSAGetLastError would give it
Public Sub plum_sock_called()
    Dim e As Long
    e = Err.LastDllError
    If e <> 0 Then sockerrno_ = e
End Sub

' #define sockerrno ((int)WSAGetLastError())
Public Function sockerrno() As Long
    sockerrno = sockerrno_
End Function

' #define poll WSAPoll
Public Function poll(pfd() As pollfd, ByVal nfds As Long, ByVal timeout As Long) As Long
    poll = ws_WSAPoll(pfd(0), nfds, timeout)
    plum_sock_called
End Function

' (VB) the byte order helpers: ntohs / ntohl of the 2 or 4 bytes at buf(offset) (a uint32_t as its bits), and
' htons / htonl into them
Public Function get_be16(buf() As Byte, ByVal offset As Long) As Long
    get_be16 = CLng(buf(offset)) * 256 + buf(offset + 1)
End Function

Public Function get_be32(buf() As Byte, ByVal offset As Long) As Long
    get_be32 = (CLng(buf(offset) And &H7F) * &H1000000) Or (CLng(buf(offset + 1)) * &H10000) Or (CLng(buf(offset + 2)) * &H100&) Or buf(offset + 3)
    If buf(offset) And &H80 Then get_be32 = get_be32 Or &H80000000
End Function

Public Sub put_be16(buf() As Byte, ByVal offset As Long, ByVal v As Long)
    buf(offset) = (v \ 256) And &HFF&
    buf(offset + 1) = v And &HFF&
End Sub

Public Sub put_be32(buf() As Byte, ByVal offset As Long, ByVal v As Long)
    buf(offset) = ((v And &HFF000000) \ &H1000000) And &HFF&
    buf(offset + 1) = (v And &HFF0000) \ &H10000
    buf(offset + 2) = (v And &HFF00&) \ &H100&
    buf(offset + 3) = v And &HFF&
End Sub
