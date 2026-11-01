Attribute VB_Name = "nettests_main"
' NetTests.exe (tools/net_test.py): WzNet's connection layer (cluster J's JT modules: tcp/*, client_connection,
' the pending writes, the zlib adapter) driven as netplay drives it, one command per run:
'   NetTests.exe <command file> <results file>
' The command file's first line is the command and its arguments (tab-separated); the results are lines of
' tab-separated fields ("result", key, value ...), with the layer's log lines ("log") and asserts ("assert").
' A connection's data is gen(seed, i); a hash is FNV-1a (32 bits) over the bytes.
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private results As Collection
Private prov As WzConnectionProvider
Private resultsPath As String

Public Sub addResult(ByVal s As String)
    If results Is Nothing Then Set results = New Collection
    results.Add s
End Sub

Private Sub res(ByVal key As String, ByVal value As Variant)
    addResult "result" & vbTab & key & vbTab & CStr(value)
End Sub

Private Function gen(ByVal seed As Long, ByVal i As Long) As Byte
    gen = (i * 31 + seed * 7 + (i \ 256) + (i \ 65536) * 3) And &HFF&
End Function

Private Function fnv(b() As Byte, ByVal n As Long) As Long
    Dim h As Long, i As Long
    h = &H811C9DC5
    For i = 0 To n - 1
        h = h Xor b(i)
        h = MUL32(h, 16777619)
    Next
    fnv = h
End Function

Private Sub setup()
    ConnectionProviderRegistry_Register 0
    Set prov = ConnectionProviderRegistry_Get(0)
    prov.initialize
    PendingWritesManagerMap_get(prov).initialize prov
End Sub

Private Sub teardown()
    Set prov = Nothing
    PendingWritesManagerMap_Shutdown
    ConnectionProviderRegistry_Shutdown
End Sub

Private Function connectSync(ByVal host As String, ByVal port As Long, ByVal timeout As Long, ec As Long) As IClientConnection
    Dim addr As IConnectionAddress
    Set addr = prov.resolveHost(host, port, ec)
    If ec <> 0 Then Exit Function
    Set connectSync = prov.openClientConnectionAny(addr, timeout, ec)
End Function

Private Function acceptOne(ByVal ls As IListenSocket, ByVal timeoutMs As Long) As IClientConnection
    Dim t0 As Long
    t0 = GetTickCount()
    Do
        Set acceptOne = ls.accept()
        If Not acceptOne Is Nothing Then Exit Function
        Sleep 1
    Loop While GetTickCount() - t0 < timeoutMs
End Function

' Writes n bytes of gen(seed) to conn in chunks of 1 .. maxChunk bytes (a fixed pattern), flushing after each;
' returns the error (0: none)
Private Function sendData(ByVal conn As IClientConnection, ByVal seed As Long, ByVal n As Long, ByVal maxChunk As Long) As Long
    Dim buf() As Byte, i As Long, pos As Long, chunk As Long, k As Long, ec As Long, raw As Long
    ReDim buf(0 To n - 1)
    For i = 0 To n - 1
        buf(i) = gen(seed, i)
    Next
    pos = 0
    Do While pos < n
        k = k + 1
        chunk = ((k * 7919) Mod maxChunk) + 1
        If chunk > n - pos Then chunk = n - pos
        conn.writeAll buf, pos, chunk, raw, ec
        If ec <> 0 Then sendData = ec: Exit Function
        ec = conn.flush(raw)
        If ec <> 0 Then sendData = ec: Exit Function
        PendingWritesManager_pollAll
        pos = pos + chunk
    Loop
End Function

' Reads n bytes from conn as netplay does (a poll group's checkConnectionsReadable, then readNoInt when readReady)
' in reads of at most maxRead bytes; returns the count read; data gets them; ec the error
Private Function recvData(ByVal conn As IClientConnection, ByVal n As Long, ByVal maxRead As Long, ByVal timeoutMs As Long, data() As Byte, ec As Long) As Long
    Dim grp As IConnectionPollGroup, got As Long, t0 As Long, r As Long, raw As Long, want As Long, ready As Long
    ReDim data(0 To n + maxRead)
    Set grp = prov.newConnectionPollGroup()
    grp.add conn
    t0 = GetTickCount()
    Do While got < n
        If GetTickCount() - t0 > timeoutMs Then Exit Do
        PendingWritesManager_pollAll
        ready = grp.checkConnectionsReadable(20, ec)
        If ec <> 0 Then Exit Do
        If ready > 0 And conn.readReady() Then
            want = maxRead
            If want > n - got Then want = n - got
            r = conn.readNoInt(data, got, want, raw, ec)
            If ec <> 0 Then Exit Do
            got = got + r
        End If
    Loop
    grp.remove conn
    recvData = got
End Function

Private Function checkData(data() As Byte, ByVal n As Long, ByVal seed As Long) As Long
    Dim i As Long
    For i = 0 To n - 1
        If data(i) <> gen(seed, i) Then checkData = i + 1: Exit Function
    Next
End Function

Private Sub errResult(ByVal key As String, ByVal ec As Long)
    res key & "_ec", ec
    If ec <> 0 Then
        res key & "_value", net_errorValue(ec)
        res key & "_category", net_errorCategory(ec)
        res key & "_categoryName", net_errorCategoryName(ec)
        res key & "_message", net_errorMessage(ec)
    End If
End Sub

' loopback <port> <bytes> <compressed> <maxChunk> <maxRead>: listen, connect, accept, gen data both ways
Private Sub cmdLoopback(a() As String)
    Dim port As Long, n As Long, compressed As Boolean, maxChunk As Long, maxRead As Long
    Dim ls As IListenSocket, client As IClientConnection, server As IClientConnection, ec As Long
    Dim data() As Byte, got As Long, t0 As Long, buf(0 To 15) As Byte, raw As Long, r As Long, i As Long
    port = CLng(a(1)): n = CLng(a(2)): compressed = (a(3) = "1"): maxChunk = CLng(a(4)): maxRead = CLng(a(5))
    setup
    Set ls = prov.openListenSocket(port, ec)
    errResult "listen", ec
    If ec <> 0 Then Exit Sub
    res "supportedIpVersions", ls.supportedIpVersions()
    res "listenType", ls.type_()
    Set client = connectSync("127.0.0.1", port, 3000, ec)
    errResult "connect", ec
    If ec <> 0 Then Exit Sub
    Set server = acceptOne(ls, 3000)
    res "accepted", Not server Is Nothing
    If server Is Nothing Then Exit Sub
    res "clientAddress", client.textAddress()
    res "serverAddress", server.textAddress()
    res "clientValid", client.isValid()
    client.useNagleAlgorithm False
    If compressed Then
        client.enableCompression
        server.enableCompression
    End If
    res "compressed", client.isCompressed() And server.isCompressed()
    t0 = GetTickCount()
    errResult "send1", sendData(client, 1, n, maxChunk)
    got = recvData(server, n, maxRead, 15000, data, ec)
    errResult "recv1", ec
    res "got1", got
    res "bad1", checkData(data, got, 1)
    errResult "send2", sendData(server, 2, n, maxChunk)
    got = recvData(client, n, maxRead, 15000, data, ec)
    errResult "recv2", ec
    res "got2", got
    res "bad2", checkData(data, got, 2)
    res "ms", GetTickCount() - t0
    errResult "status_client", client.connectionStatus()
    errResult "status_server", server.connectionStatus()
    If Not compressed Then
        ' readAll: exactly 16 bytes
        errResult "send3", sendData(client, 3, 16, 16)
        r = server.readAll(buf, 0, 16, 2000, ec)
        errResult "readAll", ec
        res "readAll_count", r
        res "readAll_bad", checkData(buf, 16, 3)
    End If
    ' the peer goes away
    client.close_
    PendingWritesManager_pollAll
    Sleep 100
    errResult "status_after_close", server.connectionStatus()
    ' (a compressed read first empties what the inflater holds: it reaches recv, and the reset, on a later call)
    For i = 1 To 3
        r = server.readNoInt(buf, 0, 16, raw, ec)
        If ec <> 0 Then Exit For
    Next
    errResult "read_after_close", ec
    server.close_
    ls.close_
    res "clientValidAfterClose", client.isValid()
    teardown
End Sub

' connect_send <host> <port> <bytes> <seed> <compressed> <maxChunk>: connect to the Python peer, (send "G",) send
' gen data, close; the peer checks it
Private Sub cmdConnectSend(a() As String)
    Dim conn As IClientConnection, ec As Long, t0 As Long
    setup
    Set conn = connectSync(a(1), CLng(a(2)), 3000, ec)
    errResult "connect", ec
    If ec <> 0 Then Exit Sub
    If a(5) = "1" Then conn.enableCompression
    errResult "send", sendData(conn, CLng(a(4)), CLng(a(3)), CLng(a(6)))
    conn.close_
    t0 = GetTickCount()
    Do While conn.isValid() And GetTickCount() - t0 < 10000   ' until the pending writes are out and it is deleted
        PendingWritesManager_pollAll
        Sleep 1
    Loop
    res "drained", Not conn.isValid()
    teardown
End Sub

' connect_recv <host> <port> <bytes> <seed> <compressed> <maxRead>: connect, send "G" (plain), maybe enable
' compression, read the peer's gen data
Private Sub cmdConnectRecv(a() As String)
    Dim conn As IClientConnection, ec As Long, data() As Byte, got As Long, g(0 To 0) As Byte, raw As Long, n As Long
    setup
    Set conn = connectSync(a(1), CLng(a(2)), 3000, ec)
    errResult "connect", ec
    If ec <> 0 Then Exit Sub
    g(0) = Asc("G")
    conn.writeAll g, 0, 1, raw, ec
    errResult "go", ec
    If a(5) = "1" Then conn.enableCompression
    n = CLng(a(3))
    got = recvData(conn, n, CLng(a(6)), 20000, data, ec)
    errResult "recv", ec
    res "got", got
    res "bad", checkData(data, got, CLng(a(4)))
    res "hash", fnv(data, got)
    conn.close_
    teardown
End Sub

' listen_recv <port> <bytes> <seed> <maxRead>: listen; the peer connects and sends gen data
Private Sub cmdListenRecv(a() As String)
    Dim ls As IListenSocket, conn As IClientConnection, ec As Long, data() As Byte, got As Long
    setup
    Set ls = prov.openListenSocket(CLng(a(1)), ec)
    errResult "listen", ec
    If ec <> 0 Then Exit Sub
    res "listening", 1
    flushResults                                 ' (the peer waits for this line)
    Set conn = acceptOne(ls, 10000)
    res "accepted", Not conn Is Nothing
    If conn Is Nothing Then Exit Sub
    res "address", conn.textAddress()
    got = recvData(conn, CLng(a(2)), CLng(a(4)), 20000, data, ec)
    errResult "recv", ec
    res "got", got
    res "bad", checkData(data, got, CLng(a(3)))
    conn.close_
    ls.close_
    teardown
End Sub

' async <host> <port> <timeoutMs>: openClientConnectionAsync with a callback
Private Sub cmdAsync(a() As String)
    Dim cb As TestOpenCallback, ok As Boolean, conn As IClientConnection, t0 As Long
    setup
    Set cb = New TestOpenCallback
    t0 = GetTickCount()
    ok = prov.openClientConnectionAsync(a(1), CLng(a(2)), CLng(a(3)), cb)
    res "returned", ok
    res "calledBeforeReturn", cb.called
    res "ms", GetTickCount() - t0
    If cb.result Is Nothing Then Exit Sub
    res "hasError", cb.result.hasError()
    errResult "error", cb.result.errorCode()
    res "errorString", cb.result.errorString()
    res "timed_out", net_errorEqualsErrc(cb.result.errorCode(), errc_timed_out)
    res "connection_refused", net_errorEqualsErrc(cb.result.errorCode(), errc_connection_refused)
    If Not cb.result.hasError() Then
        Set conn = cb.result.open_socket_release()
        res "address", conn.textAddress()
        conn.close_
    End If
    Set cb = Nothing
    teardown
End Sub

' resolve <host> <port>
Private Sub cmdResolve(a() As String)
    Dim addr As IConnectionAddress, ec As Long, s As String
    setup
    Set addr = prov.resolveHost(a(1), CLng(a(2)), ec)
    errResult "resolve", ec
    If ec = 0 Then
        s = addr.toString(ec)
        errResult "toString", ec
        res "address", s
    End If
    Set addr = Nothing
    teardown
End Sub

' readall_timeout <port>: readAll on a connection with nothing to read
Private Sub cmdReadAllTimeout(a() As String)
    Dim ls As IListenSocket, client As IClientConnection, server As IClientConnection, ec As Long, buf(0 To 9) As Byte, t0 As Long
    setup
    Set ls = prov.openListenSocket(CLng(a(1)), ec)
    Set client = connectSync("127.0.0.1", CLng(a(1)), 3000, ec)
    Set server = acceptOne(ls, 3000)
    t0 = GetTickCount()
    server.readAll buf, 0, 10, 300, ec
    res "ms", GetTickCount() - t0
    errResult "readAll", ec
    res "timed_out", net_errorEqualsErrc(ec, errc_timed_out)
    client.close_
    server.close_
    ls.close_
    teardown
End Sub

' pending_full <host> <port> <bytes> <seed> <chunk>: queue everything at once while the peer is not reading, then
' let the pending writes drain
Private Sub cmdPendingFull(a() As String)
    Dim conn As IClientConnection, ec As Long, buf() As Byte, n As Long, chunk As Long, pos As Long, i As Long
    Dim t0 As Long, raw As Long, passes As Long
    setup
    Set conn = connectSync(a(1), CLng(a(2)), 3000, ec)
    errResult "connect", ec
    If ec <> 0 Then Exit Sub
    n = CLng(a(3)): chunk = CLng(a(5))
    ReDim buf(0 To n - 1)
    For i = 0 To n - 1
        buf(i) = gen(CLng(a(4)), i)
    Next
    t0 = GetTickCount()
    Do While pos < n
        If chunk > n - pos Then chunk = n - pos
        conn.writeAll buf, pos, chunk, raw, ec
        If ec <> 0 Then Exit Do
        pos = pos + chunk
    Loop
    errResult "write", ec
    res "queued_ms", GetTickCount() - t0
    res "queued", pos
    conn.close_
    res "validAfterClose", conn.isValid()      ' still pending: deleted once written
    t0 = GetTickCount()
    Do While conn.isValid() And GetTickCount() - t0 < 30000
        PendingWritesManager_pollAll
        passes = passes + 1
        Sleep 2
    Loop
    res "drained", Not conn.isValid()
    res "drain_ms", GetTickCount() - t0
    res "passes", passes
    teardown
End Sub

' errc: the error codes' messages and comparisons
Private Sub cmdErrc(a() As String)
    Dim v As Variant
    For Each v In Array(0, 9, 14, 22, 108, 138, 10004, 10009, 10035, 10036, 10054, 10060, 10061, 10065, 11001, 11004, 12345)
        res "gs_" & v, net_errorMessage(make_network_error_code(v))
    Next
    For Each v In Array(4, 11, 106, 107, 108, 138, 140, 999)
        res "generic_" & v, net_errorMessage(make_generic_error_code(v))
    Next
    For Each v In Array(11001, 11002, 11003, 10109)
        res "gai_" & v, net_errorMessage(make_getaddrinfo_error_code(v))
    Next
    For Each v In Array(2, -2, -3, -4, -5)
        res "zlib_" & v, net_errorMessage(make_zlib_error_code(v))
    Next
    res "eq_10060_timed_out", net_errorEqualsErrc(make_network_error_code(10060), errc_timed_out)
    res "eq_138_timed_out", net_errorEqualsErrc(make_network_error_code(138), errc_timed_out)
    res "eq_10035_would_block", net_errorEqualsErrc(make_network_error_code(10035), errc_operation_would_block)
    res "eq_10035_try_again", net_errorEqualsErrc(make_network_error_code(10035), errc_resource_unavailable_try_again)
    res "eq_10054_reset", net_errorEqualsErrc(make_network_error_code(10054), errc_connection_reset)
    res "eq_108_reset", net_errorEqualsErrc(make_network_error_code(108), errc_connection_reset)
    res "eq_10004_interrupted", net_errorEqualsErrc(make_network_error_code(10004), errc_interrupted)
    res "eq_generic_106_aborted", net_errorEqualsErrc(make_generic_error_code(106), errc_connection_aborted)
    res "eq_gai_11001_timed_out", net_errorEqualsErrc(make_getaddrinfo_error_code(10060), errc_timed_out)
    res "eq_zero", net_errorEqualsErrc(0, 0)
    res "value_zlib_-3", net_errorValue(make_zlib_error_code(-3))
    res "value_0_is_error", make_network_error_code(0) <> 0
End Sub

' byteorder / ip helpers
Private Sub cmdHelpers(a() As String)
    Dim b(0 To 3) As Byte, w As Long, v As Variant
    res "htonl", Hex$(wz_htonl(&H12345678))
    res "htonl2", Hex$(wz_htonl(&H80FF0102))
    res "ntohl", Hex$(wz_ntohl(&H78563412))
    res "htons", Hex$(wz_htons(&H1234))
    res "htons2", Hex$(wz_htons(&HFF80&))
    wz_htons_store_unaligned b, 1, &HABCD&
    res "store", Hex$(b(1)) & " " & Hex$(b(2))
    wz_ntohs_load_unaligned w, b, 1
    res "load", Hex$(w)
    For Each v In Array("10.1.2.3", "172.16.0.1", "172.32.0.1", "192.168.1.1", "169.254.3.4", "127.0.0.1", "8.8.8.8", _
                         "::1", "fe80::1", "fd00::5", "2001:db8::1", "::ffff:192.168.0.9", "::ffff:8.8.4.4", "nonsense", "")
        res "class_" & v, GetIPClassification(CStr(v))
        res "format_" & v, GetIPFormat(CStr(v))
    Next
    res "class_null", GetIPClassification(vbNullString)
    res "format_names", IPFormat_to_string(0) & "," & IPFormat_to_string(1) & "," & IPFormat_to_string(2) & "," & IPFormat_to_string(9)
End Sub

Private Sub flushResults()
    Dim f As Integer, s As Variant
    If results Is Nothing Then Exit Sub
    f = FreeFile
    Open resultsPath For Output As #f
    For Each s In results
        Print #f, s
    Next
    Close #f
End Sub

Sub Main()
    Dim args() As String, f As Integer, line_ As String, a() As String
    On Error GoTo fail
    args = Split(Command$, " ")
    resultsPath = Replace(args(1), """", "")
    f = FreeFile
    Open Replace(args(0), """", "") For Input As #f
    Line Input #f, line_
    Close #f
    a = Split(line_, vbTab)
    Select Case a(0)
    Case "loopback": cmdLoopback a
    Case "connect_send": cmdConnectSend a
    Case "connect_recv": cmdConnectRecv a
    Case "listen_recv": cmdListenRecv a
    Case "async": cmdAsync a
    Case "resolve": cmdResolve a
    Case "readall_timeout": cmdReadAllTimeout a
    Case "pending_full": cmdPendingFull a
    Case "errc": cmdErrc a
    Case "helpers": cmdHelpers a
    End Select
    res "asserts", nettests_stubs.asserts
    res "done", 1
    flushResults
    Exit Sub
fail:
    res "vb_error", Err.Number & " " & Err.Description
    res "asserts", nettests_stubs.asserts
    flushResults
End Sub
