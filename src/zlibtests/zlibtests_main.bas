Attribute VB_Name = "zlibtests_main"
' (VB) ZlibTests: src/shared/miniz's streams (miniz_tdefl_c, miniz_tinfl_c, miniz_c) driven by tools/zlib_test.py.
' ZlibTests.exe <commands file> <results file>: each line of the commands file (tab-separated) gives one line of
' results (tab-separated: the command's first field, then "ok" or "err" and numbers):
'   def <in> <out> <level> <strategy> <window bits> <in piece> <out chunk> <flush>
'       mz_deflateInit2 (MZ_DEFLATED, mem_level 9); the input in pieces of <in piece> bytes (0: the sizes listed in
'       <in>.sizes, one a line), each deflated with output room of <out chunk> bytes a call and <flush> after it
'       (0 none, 1 partial, 2 sync, 3 full; 9: zlib_compression_adapter.cpp's pattern: compress() with the piece
'       and room for its size + 20, again while the room ran out, then flushCompressionStream()'s Z_PARTIAL_FLUSH
'       with no input and room for 1000, again while that ran out); then Z_FINISH to the end (unless <flush> is
'       9 and <out chunk> is 0: the network's stream never ends). Writes <out>, and <out>.marks: the output's
'       length after each piece. Results: ok, total_in, total_out, adler (unsigned), the last status
'   inf <in> <out> <window bits> <in piece> <out chunk> <flush>
'       mz_inflateInit2; the input in pieces (0: <in>.sizes), each inflated with room of <out chunk> bytes a call,
'       again while the room ran out or input is left (until an error, Z_STREAM_END, or Z_BUF_ERROR: no
'       progress). Writes <out> and <out>.marks. Results: ok, the last status, total_in, total_out, adler
'   adapt <in> <out> <max size> <received>
'       the network's receive (client_connection.cpp's readNoInt over zlib_compression_adapter.cpp): while input is
'       needed, "receive" the next part of <in> into the adapter's own buffer (<received> 0: max size + 1000 bytes,
'       its buffer's size; 1: the next of <in>.sizes, at most that) and make it the whole input; one inflate with
'       Z_NO_FLUSH and room for <max size>; with room left, input is needed next time (the adapter asserts all of it
'       was taken: unconsumed bytes are dropped, as there). Writes <out> and <out>.calls (each call: produced,
'       avail_in, avail_out, status). Results: ok, calls, calls with room and input left (the assert), last status
'   cmp2 <in> <out> <level>          mz_compress2 one-shot. Results: ok, status, the compressed size
'   speed <in> <level> <reps>        deflate (Z_FINISH, one call) and inflate the file <reps> times. Results: ok,
'                                    deflate MB/s, inflate MB/s (of the uncompressed size)
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal src As Long, ByVal n As Long)

Private outBuf() As Byte
Private outLen As Long
Private outCap As Long
Private marks As String

Private Function readFile(ByVal path As String, b() As Byte) As Long
    Dim f As Integer, n As Long
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    Else
        Erase b
    End If
    Close #f
    readFile = n
End Function

Private Sub writeBytes(ByVal path As String, b() As Byte, ByVal n As Long)
    Dim f As Integer, c() As Byte
    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    If n > 0 Then
        ReDim c(0 To n - 1)
        CopyMemory VarPtr(c(0)), VarPtr(b(0)), n
        Put #f, , c
    End If
    Close #f
End Sub

Private Sub writeText(ByVal path As String, ByVal s As String)
    Dim f As Integer
    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Output As #f
    Print #f, s;
    Close #f
End Sub

' the piece sizes: n bytes in pieces of piece bytes, or those listed in path (one a line)
Private Function pieceSizes(ByVal path As String, ByVal n As Long, ByVal piece As Long, sizes() As Long) As Long
    Dim count As Long, f As Integer, ln As String
    count = 0
    If piece > 0 Then
        ReDim sizes(0 To n \ piece + 1)
        Do While n > 0
            If n < piece Then sizes(count) = n Else sizes(count) = piece
            n = n - sizes(count): count = count + 1
        Loop
    Else
        ReDim sizes(0 To 1023)
        f = FreeFile
        Open path For Input As #f
        Do While Not EOF(f)
            Line Input #f, ln
            If Len(Trim$(ln)) > 0 Then
                If count > UBound(sizes) Then ReDim Preserve sizes(0 To count * 2)
                sizes(count) = CLng(ln): count = count + 1
            End If
        Loop
        Close #f
    End If
    pieceSizes = count
End Function

' room for n more bytes at outLen
Private Sub ensureRoom(ByVal n As Long)
    If outCap < outLen + n Then
        If outCap = 0 Then outCap = 4096
        Do While outCap < outLen + n
            outCap = outCap * 2
        Loop
        ReDim Preserve outBuf(0 To outCap - 1)
    End If
End Sub

Private Function unsignedText(ByVal x As Long) As String
    If x < 0 Then unsignedText = CStr(CDbl(x) + 4294967296#) Else unsignedText = CStr(x)
End Function

Private Function cmdDeflate(f() As String) As String
    Dim inData() As Byte, n As Long, level As Long, strategy As Long, wbits As Long, piece As Long, chunk As Long, flushMode As Long
    Dim strm As z_stream, sizes() As Long, count As Long, i As Long, pos As Long, status As Long, noInput() As Byte, room As Long
    n = readFile(f(1), inData)
    level = CLng(f(3)): strategy = CLng(f(4)): wbits = CLng(f(5)): piece = CLng(f(6)): chunk = CLng(f(7)): flushMode = CLng(f(8))
    count = pieceSizes(f(1) & ".sizes", n, piece, sizes)
    outLen = 0: outCap = 0: Erase outBuf: marks = ""
    status = mz_deflateInit2(strm, level, MZ_DEFLATED, wbits, 9, strategy)
    If status <> Z_OK Then
        cmdDeflate = "err" & vbTab & "init " & status
        Exit Function
    End If
    pos = 0
    For i = 0 To count - 1
        If flushMode = 9 Then
            ' ZlibCompressionAdapter::compress(src, size)
            strm.next_in = pos: strm.avail_in = sizes(i)
            Do
                room = sizes(i) + 20
                ensureRoom room
                strm.next_out = outLen: strm.avail_out = room
                status = deflate(strm, inData, outBuf, Z_NO_FLUSH)
                If status = Z_STREAM_ERROR Then
                    cmdDeflate = "err" & vbTab & "compress " & i
                    Exit Function
                End If
                outLen = outLen + room - strm.avail_out
            Loop While strm.avail_out = 0
            ' ZlibCompressionAdapter::flushCompressionStream()
            Do
                strm.avail_in = 0                       ' next_in = nullptr
                room = 1000
                ensureRoom room
                strm.next_out = outLen: strm.avail_out = room
                status = deflate(strm, noInput, outBuf, Z_PARTIAL_FLUSH)
                If status = Z_STREAM_ERROR Then
                    cmdDeflate = "err" & vbTab & "flush " & i
                    Exit Function
                End If
                outLen = outLen + room - strm.avail_out
            Loop While strm.avail_out = 0
        Else
            strm.next_in = pos: strm.avail_in = sizes(i)
            Do
                ensureRoom chunk
                strm.next_out = outLen: strm.avail_out = chunk
                status = deflate(strm, inData, outBuf, flushMode)
                outLen = outLen + chunk - strm.avail_out
                If status < 0 And status <> Z_BUF_ERROR Then
                    cmdDeflate = "err" & vbTab & "deflate " & status & " piece " & i
                    Exit Function
                End If
                If strm.avail_in = 0 Then
                    If flushMode = Z_NO_FLUSH Then Exit Do
                    If strm.avail_out <> 0 Then Exit Do
                End If
            Loop
        End If
        pos = pos + sizes(i)
        marks = marks & outLen & vbCrLf
    Next
    If Not (flushMode = 9 And chunk = 0) Then
        If chunk <= 0 Then chunk = 4096
        strm.avail_in = 0
        Do
            ensureRoom chunk
            strm.next_out = outLen: strm.avail_out = chunk
            status = deflate(strm, noInput, outBuf, Z_FINISH)
            outLen = outLen + chunk - strm.avail_out
            If status = Z_STREAM_END Then Exit Do
            If status < 0 And status <> Z_BUF_ERROR Then
                cmdDeflate = "err" & vbTab & "finish " & status
                Exit Function
            End If
        Loop
    End If
    writeBytes f(2), outBuf, outLen
    writeText f(2) & ".marks", marks
    cmdDeflate = "ok" & vbTab & strm.total_in & vbTab & strm.total_out & vbTab & unsignedText(strm.adler) & vbTab & status
    deflateEnd strm
End Function

Private Function cmdInflate(f() As String) As String
    Dim inData() As Byte, n As Long, wbits As Long, piece As Long, chunk As Long, flushMode As Long
    Dim strm As z_stream, sizes() As Long, count As Long, i As Long, pos As Long, status As Long
    n = readFile(f(1), inData)
    wbits = CLng(f(3)): piece = CLng(f(4)): chunk = CLng(f(5)): flushMode = CLng(f(6))
    count = pieceSizes(f(1) & ".sizes", n, piece, sizes)
    outLen = 0: outCap = 0: Erase outBuf: marks = ""
    status = mz_inflateInit2(strm, wbits)
    If status <> Z_OK Then
        cmdInflate = "err" & vbTab & "init " & status
        Exit Function
    End If
    pos = 0
    For i = 0 To count - 1
        strm.next_in = pos: strm.avail_in = sizes(i)
        Do
            ensureRoom chunk
            strm.next_out = outLen: strm.avail_out = chunk
            status = inflate(strm, inData, outBuf, flushMode)
            outLen = outLen + chunk - strm.avail_out
            If status < 0 Or status = Z_STREAM_END Then Exit Do
            If (strm.avail_in = 0) And (strm.avail_out <> 0) Then Exit Do
        Loop
        pos = pos + sizes(i)
        marks = marks & outLen & vbCrLf
        If status < 0 And status <> Z_BUF_ERROR Then Exit For
        If status = Z_STREAM_END Then Exit For
    Next
    writeBytes f(2), outBuf, outLen
    writeText f(2) & ".marks", marks
    cmdInflate = "ok" & vbTab & status & vbTab & strm.total_in & vbTab & strm.total_out & vbTab & unsignedText(strm.adler)
    inflateEnd strm
End Function

Private Function cmdAdapter(f() As String) As String
    Dim inData() As Byte, n As Long, maxSize As Long, received As Long, strm As z_stream, status As Long
    Dim sizes() As Long, count As Long, segIdx As Long, segLeft As Long, recvPos As Long, size As Long
    Dim inflateInBuf() As Byte, needInput As Boolean, calls As Long, violations As Long, produced As Long, log_ As String
    n = readFile(f(1), inData)
    maxSize = CLng(f(3)): received = CLng(f(4))
    If received = 1 Then count = pieceSizes(f(1) & ".sizes", n, 0, sizes)
    outLen = 0: outCap = 0: Erase outBuf
    status = inflateInit(strm)
    If status <> Z_OK Then
        cmdAdapter = "err" & vbTab & "init " & status
        Exit Function
    End If
    ReDim inflateInBuf(0 To maxSize + 1000 - 1)         ' decompressInBuf.resize(max_size + 1000)
    needInput = True: recvPos = 0: segIdx = 0: segLeft = 0
    Do
        If needInput Then
            If recvPos >= n Then Exit Do
            If received = 1 Then
                If segLeft = 0 Then segLeft = sizes(segIdx): segIdx = segIdx + 1
                size = segLeft
                If size > maxSize + 1000 Then size = maxSize + 1000
                segLeft = segLeft - size
            Else
                size = n - recvPos
                If size > maxSize + 1000 Then size = maxSize + 1000
            End If
            If size > 0 Then CopyMemory VarPtr(inflateInBuf(0)), VarPtr(inData(recvPos)), size
            recvPos = recvPos + size
            strm.next_in = 0: strm.avail_in = size     ' resetDecompressionStreamInputSize(received)
            needInput = False
        End If
        ensureRoom maxSize
        strm.next_out = outLen: strm.avail_out = maxSize
        status = inflate(strm, inflateInBuf, outBuf, Z_NO_FLUSH)
        If status = Z_STREAM_ERROR Then Exit Do
        produced = maxSize - strm.avail_out
        outLen = outLen + produced
        calls = calls + 1
        log_ = log_ & produced & " " & strm.avail_in & " " & strm.avail_out & " " & status & vbCrLf
        If strm.avail_out <> 0 Then
            needInput = True
            If strm.avail_in <> 0 Then violations = violations + 1  ' ASSERT(decompressionStreamConsumedAllInput())
        End If
        If status = Z_STREAM_END Or status = Z_DATA_ERROR Then Exit Do
    Loop
    writeBytes f(2), outBuf, outLen
    writeText f(2) & ".calls", log_
    cmdAdapter = "ok" & vbTab & calls & vbTab & violations & vbTab & status
    inflateEnd strm
End Function

Private Function cmdCompress2(f() As String) As String
    Dim inData() As Byte, n As Long, dest() As Byte, destLen As Long, status As Long
    n = readFile(f(1), inData)
    destLen = CLng(mz_compressBound(n))
    ReDim dest(0 To destLen - 1)
    status = mz_compress2(dest, destLen, inData, n, CLng(f(3)))
    If status = Z_OK Then writeBytes f(2), dest, destLen
    cmdCompress2 = "ok" & vbTab & status & vbTab & destLen
End Function

Private Function cmdSpeed(f() As String) As String
    Dim inData() As Byte, n As Long, level As Long, reps As Long, r As Long, strm As z_stream, t0 As Double, tDef As Double, tInf As Double
    Dim comp() As Byte, compLen As Long, back() As Byte, status As Long
    n = readFile(f(1), inData)
    level = CLng(f(2)): reps = CLng(f(3))
    ReDim comp(0 To CLng(mz_compressBound(n)) - 1)
    ReDim back(0 To n)
    For r = 1 To reps
        t0 = Timer
        mz_deflateInit2 strm, level, MZ_DEFLATED, MZ_DEFAULT_WINDOW_BITS, 9, MZ_DEFAULT_STRATEGY
        strm.next_in = 0: strm.avail_in = n: strm.next_out = 0: strm.avail_out = UBound(comp) + 1
        status = deflate(strm, inData, comp, Z_FINISH)
        compLen = strm.total_out
        deflateEnd strm
        tDef = tDef + (Timer - t0)
        t0 = Timer
        inflateInit strm
        strm.next_in = 0: strm.avail_in = compLen: strm.next_out = 0: strm.avail_out = n + 1
        status = inflate(strm, comp, back, Z_FINISH)
        inflateEnd strm
        tInf = tInf + (Timer - t0)
    Next
    If tDef <= 0 Then tDef = 0.001
    If tInf <= 0 Then tInf = 0.001
    cmdSpeed = "ok" & vbTab & Format$(n * reps / tDef / 1048576#, "0.00") & vbTab & Format$(n * reps / tInf / 1048576#, "0.00")
End Function

Public Sub Main()
    Dim args As String, cmdPath As String, resPath As String, fin As Integer, ln As String, f() As String, res As String, results As String
    args = Command$
    cmdPath = Trim$(Left$(args, InStr(args, " ") - 1)): resPath = Trim$(Mid$(args, InStr(args, " ") + 1))
    If Left$(cmdPath, 1) = """" Then cmdPath = Mid$(cmdPath, 2, Len(cmdPath) - 2)
    If Left$(resPath, 1) = """" Then resPath = Mid$(resPath, 2, Len(resPath) - 2)
    fin = FreeFile
    Open cmdPath For Input As #fin
    Do While Not EOF(fin)
        Line Input #fin, ln
        If Len(ln) > 0 Then
            f = Split(ln, vbTab)
            Select Case f(0)
            Case "def": res = cmdDeflate(f)
            Case "inf": res = cmdInflate(f)
            Case "adapt": res = cmdAdapter(f)
            Case "cmp2": res = cmdCompress2(f)
            Case "speed": res = cmdSpeed(f)
            Case Else: res = "err" & vbTab & "unknown command"
            End Select
            results = results & f(0) & vbTab & res & vbCrLf
        End If
    Loop
    Close #fin
    writeText resPath, results
End Sub
