Attribute VB_Name = "netlog_c"
' lib/netplay/netlog.cpp, netlog.h: Logging for debug only (the write dir's logs/netplay-<date>_<time>.log,
' opened by NETinit, closed by NETshutdown) and the packet counters.
' getAscTime (wztime.cpp: put_time's "%c" in the classic locale) is the Windows build's: "%m/%d/%y %H:%M:%S".
Option Explicit

Private Const NUM_GAME_PACKETS As Long = 256

Private pFileHandle As Long                 ' PHYSFS_file *
Private packetcount(0 To 1, 0 To NUM_GAME_PACKETS - 1) As Long  ' uint32_t
Private packetsize(0 To 1, 0 To NUM_GAME_PACKETS - 1) As Long   ' uint32_t
Private lastframe As Long                   ' NETlogEntry's static UDWORD lastframe

Private Function getAscTime(tmVal As tm) As String
    getAscTime = strftime("%m/%d/%y %H:%M:%S", tmVal)
End Function

' WZ_PHYSFS_writeBytes(pFileHandle, buf, strlen(buf))
Private Sub writeStr(ByVal s As String)
    Dim b() As Byte, n As Long
    n = utf8_encode(s, b)
    If n > 0 Then WZ_PHYSFS_writeBytes pFileHandle, b, 0, n
End Sub

Public Function NETstartLogging() As Boolean
    Dim aclock As Double, buf As String, filename As String, i As Long, newtime As tm

    For i = 0 To NUM_GAME_PACKETS - 1
        packetcount(0, i) = 0
        packetsize(0, i) = 0
        packetcount(1, i) = 0
        packetsize(1, i) = 0
    Next

    aclock = time_()                        ' Get time in seconds
    newtime = getLocalTime(aclock)          ' Convert time to struct

    filename = Left$(va("logs/netplay-%04d%02d%02d_%02d%02d%02d.log", newtime.tm_year + 1900, newtime.tm_mon + 1, newtime.tm_mday, newtime.tm_hour, newtime.tm_min, newtime.tm_sec), 255)
    pFileHandle = PHYSFS_openWrite(filename)    ' open the file
    If pFileHandle = 0 Then
        debug_ 65, LOG_ERROR, "NETstartLogging", "Could not create net log %s: %s", filename, WZ_PHYSFS_getLastError()
        NETstartLogging = False
        Exit Function
    End If
    buf = Left$(va("NETPLAY log: %s" & vbLf, getAscTime(newtime)), 255)
    writeStr buf
    NETstartLogging = True
End Function

Public Function NETstopLogging() As Boolean
    Const dash_line As String = "-----------------------------------------------------------" & vbLf
    Dim buf As String, i As Long, pname As String, IPAddress As String
    Dim totalBytessent As Long, totalBytesrecv As Long, totalPacketsent As Long, totalPacketrecv As Long   ' UDWORD

    If pFileHandle = 0 Then
        NETstopLogging = False
        Exit Function
    End If

    ' Output stats
    For i = 0 To NUM_GAME_PACKETS - 1
        If messageTypeToString(i) = "(UNUSED)" Then
            GoTo continue_
        Else
            buf = Left$(va("%-24s:" & vbTab & " received %u times, %u bytes; sent %u times, %u bytes" & vbLf, messageTypeToString(i), _
                packetcount(1, i), packetsize(1, i), packetcount(0, i), packetsize(0, i)), 255)
        End If
        writeStr buf
        totalBytessent = totalBytessent + packetsize(0, i)
        totalBytesrecv = totalBytesrecv + packetsize(1, i)
        totalPacketsent = totalPacketsent + packetcount(0, i)
        totalPacketrecv = totalPacketrecv + packetcount(1, i)
continue_:
    Next
    writeStr Left$(va("== Total bytes sent %u, Total bytes received %u ==" & vbLf, totalBytessent, totalBytesrecv), 255)
    writeStr Left$(va("== Total packets sent %u, recv %u ==" & vbLf, totalPacketsent, totalPacketrecv), 255)
    writeStr vbLf & "-Sync statistics -" & vbLf
    writeStr dash_line
    writeStr Left$(va("joins: %u, kicks: %u, drops: %u, left %u" & vbLf, sync_counter.joins, sync_counter.kicks, sync_counter.drops, sync_counter.left_), 255)
    writeStr Left$(va("banned: %u, cantjoin: %u, rejected: %u" & vbLf, sync_counter.banned, sync_counter.cantjoin, sync_counter.rejected), 255)
    If sync_counter.banned <> 0 Then
        If NETgetIPBanList_size() > 0 Then
            writeStr "Banned list:" & vbLf
            For i = 0 To NETgetIPBanList_size() - 1
                NETgetIPBanList_item i, pname, IPAddress
                If IPAddress <> "" Then
                    writeStr Left$(va("player %s, IP: %s" & vbLf, pname, IPAddress), 255)
                End If
            Next
        End If
    End If
    writeStr dash_line
    writeStr dash_line

    If PHYSFS_close(pFileHandle) = 0 Then
        debug_ 138, LOG_ERROR, "NETstopLogging", "Could not close net log: %s", WZ_PHYSFS_getLastError()
        NETstopLogging = False
        Exit Function
    End If
    pFileHandle = 0

    NETstopLogging = True
End Function

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

Public Function NETlogEntry(ByVal str_ As String, ByVal a As Long, ByVal b As Long) As Boolean
    Const star_line As String = "************************************************************" & vbLf
    Const dash_line As String = "-----------------------------------------------------------" & vbLf
    Dim frame As Long, aclock As Double, buf As String, newtime As tm

    frame = frameGetFrameNumber()

    If pFileHandle = 0 Then
        NETlogEntry = False
        Exit Function
    End If

    aclock = time_()                        ' Get time in seconds

    If Not getLocalTimeOpt(newtime, aclock) Or pFileHandle = 0 Then  ' Convert time to struct
        debug_ 176, LOG_ERROR, "NETlogEntry", "Fatal error averted in NETlog"
        NETlogEntry = False
        Exit Function
    End If

    ' check to see if a new frame.
    If frame <> lastframe Then
        lastframe = frame
        writeStr dash_line
    End If

    If ULT(a, NUM_GAME_PACKETS) Then
        ' replace common msgs with txt descriptions
        buf = va("%s # %s " & vbTab & ": %s " & vbTab & ":%d" & vbLf, getAscTime(newtime), str_, messageTypeToString(a), b)
    ElseIf a = SYNC_FLAG Then
        buf = va("%s # %s " & vbTab & ": %d " & vbTab & "(Sync)" & vbLf, getAscTime(newtime), str_, b)
    Else
        buf = va("%s # %s " & vbTab & ":%d " & vbTab & vbTab & vbTab & ":%d" & vbLf, getAscTime(newtime), str_, a, b)
    End If
    buf = Left$(buf, 255)

    If a = NET_PLAYER_LEAVING Or a = NET_PLAYER_DROPPED Then
        ' Write a starry line above NET_LEAVING messages
        writeStr star_line
        writeStr buf
        writeStr star_line
    Else
        writeStr buf
    End If

    PHYSFS_flush pFileHandle
    NETlogEntry = True
End Function
