Attribute VB_Name = "httptests_main"
' HttpTests.exe (tools/http_test.py): WzNet's urltransfer_c (curl's part of urlrequest: HTTP/1.1 over Winsock and
' Schannel) driven as urlrequest_c drives it, one transfer per run:
'   HttpTests.exe <command file> <results file>
' The command file's lines are "key<TAB>value": method, url, protocol (0 IP_ANY, 1 IPv4, 2 IPv6), timeout
' (CURLOPT_CONNECTTIMEOUT_MS), follow (0/1), maxsize, ua, body (hex; absent: no body), header (one line each),
' limit (how long the run may poll, ms). The results are lines of tab-separated fields: "result", key, value; the
' response's header lines ("header", line); the transfer's log lines ("log") and asserts ("assert").
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private results As Collection
Private resultsPath As String

Public Sub addResult(ByVal s As String)
    If results Is Nothing Then Set results = New Collection
    results.Add s
End Sub

Private Sub res(ByVal key As String, ByVal value As Variant)
    addResult "result" & vbTab & key & vbTab & CStr(value)
End Sub

Private Function toHex(b() As Byte, ByVal n As Long) As String
    Dim s As String, k As Long
    s = String$(n * 2, "0")
    For k = 0 To n - 1
        Mid$(s, k * 2 + 1, 2) = Right$("0" & Hex$(b(k)), 2)
    Next
    toHex = s
End Function

Private Function fromHex(ByVal s As String, b() As Byte) As Long
    Dim n As Long, k As Long
    n = Len(s) \ 2
    If n > 0 Then
        ReDim b(0 To n - 1)
        For k = 0 To n - 1
            b(k) = CByte("&H" & Mid$(s, k * 2 + 1, 2))
        Next
    End If
    fromHex = n
End Function

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

Private Sub runTransfer(ByVal cmdPath As String)
    Dim f As Integer, line_ As String, p As Long, key As String, value As String
    Dim method As String, url As String, protocol As Long, timeout As Long, follow As Boolean, maxSize As Double
    Dim ua As String, body() As Byte, bodyLen As Long, headers As String, limit As Long
    Dim h As Long, t0 As Long, done As Boolean, code As Long, text As String, status As Long, ip As String
    Dim retryAfter As Double, resp() As Byte, n As Long, lines() As String, k As Long, polls As Long
    method = "GET"
    bodyLen = -1
    follow = True
    limit = 120000
    f = FreeFile
    Open cmdPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, line_
        p = InStr(line_, vbTab)
        If p > 0 Then
            key = Left$(line_, p - 1)
            value = Mid$(line_, p + 1)
            Select Case key
            Case "method": method = value
            Case "url": url = value
            Case "protocol": protocol = CLng(value)
            Case "timeout": timeout = CLng(value)
            Case "follow": follow = (value = "1")
            Case "maxsize": maxSize = CDbl(value)
            Case "ua": ua = value
            Case "body": bodyLen = fromHex(value, body)
            Case "header": headers = headers & value & vbCrLf
            Case "limit": limit = CLng(value)
            End Select
        End If
    Loop
    Close #f
    t0 = GetTickCount()
    h = urlTransferStart(method, url, headers, body, bodyLen, protocol, timeout, follow, ua, maxSize)
    res "handle", h
    Do
        done = urlTransferPoll(h, 50)
        polls = polls + 1
        If done Then Exit Do
        If GetTickCount() - t0 > limit Then Exit Do
    Loop
    res "done", IIf(done, 1, 0)
    res "elapsed", GetTickCount() - t0
    res "polls", polls
    code = urlTransferResult(h, text, status, ip, retryAfter)
    res "code", code
    res "text", text
    res "status", status
    res "ip", ip
    res "retryafter", retryAfter
    lines = Split(urlTransferHeaders(h), vbLf)
    For k = 0 To UBound(lines)
        If Len(lines(k)) > 0 Then addResult "header" & vbTab & lines(k)
    Next
    n = urlTransferBody(h, resp)
    res "bodylen", n
    If n > 0 And n <= 1000000 Then res "body", toHex(resp, n)
    urlTransferClose h
    res "version", urlTransferVersion()
End Sub

Sub Main()
    Dim args() As String
    On Error GoTo fail
    args = Split(Command$, " ")
    resultsPath = Replace(args(1), """", "")
    runTransfer Replace(args(0), """", "")
    res "asserts", nettests_stubs.asserts
    res "finished", 1
    flushResults
    Exit Sub
fail:
    res "vb_error", Err.Number & " " & Err.Description
    flushResults
End Sub
