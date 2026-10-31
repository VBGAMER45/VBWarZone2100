Attribute VB_Name = "opustests_main"
' OpusTests.exe (VB-only, not part of the game; cluster K): decodes Ogg Opus files with WzSound's opusfile
' port (opusfile_c, libogg's ogg_framing_c) over libopus's port (O2's celt_*, O3's silk_* and opus_*), all
' compiled in, for tools/opus_test.py, which compares the samples with ffmpeg's libopus decoder.
'
'   OpusTests <list file> <output folder> [noseek]
'
' For the n-th file of the list (0-based, one path per line):
'   <n>.f32    op_read_float's samples (all channels, float32, interleaved)
'   <n>.s16    op_read_stereo's samples (16-bit stereo, dithered: what the game plays)
'   <n>.f2     op_read_float_stereo's samples (float32 stereo), for files of more than 2 channels
'   <n>.wz     WZOpusDecoder (fromFilename, a PhysFS stream) decoded as openal_track's sound_fillNBuffers
'              does (16 KiB buffers, until one comes back empty)
'   <n>.wzm    the same with bufferEntireStream (the memory buffer callbacks)
'   <n>.seekK.f32  4800 samples per channel read after op_pcm_seek to target K (unless noseek)
' and a line in <output folder>\results.txt: "n" then TAB-separated key=value fields (channels, pcm_total,
' links, pre_skip, gain, holes, the sample counts, the decode times in ms, the seeks' targets, returns and
' op_pcm_tell); a VB error: "n<TAB>error=<number> <description> at <stage>". The last line is "done".
Option Explicit

Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private qpcFreq As Currency
Private stage As String

Private Function GetArgs(argv() As String) As Long
    Dim p As Long, i As Long, sp As Long, n As Long, argc As Long
    p = CommandLineToArgvW(GetCommandLineW(), argc)
    ReDim argv(0 To argc)
    For i = 0 To argc - 1
        CopyMemory sp, ByVal p + i * 4, 4
        n = lstrlenW(sp)
        argv(i) = String$(n, 0)
        If n > 0 Then CopyMemory ByVal StrPtr(argv(i)), ByVal sp, n * 2
    Next
    LocalFree p
    GetArgs = argc
End Function

Private Function ReadBytes(ByVal path As String, b() As Byte) As Long
    Dim f As Integer, n As Long
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    Else
        ReDim b(0 To 0)
    End If
    Close #f
    ReadBytes = n
End Function

Private Function OutOpen(ByVal path As String) As Long
    OutOpen = CreateFileW(StrPtr(path), &H40000000, 0, 0, 2, &H80, 0)
End Function

Private Function Ms(ByVal t0 As Currency, ByVal t1 As Currency) As Double
    Ms = (t1 - t0) * 1000# / qpcFreq
End Function

Private Sub AddField(line_ As String, ByVal k As String, ByVal v As Variant)
    line_ = line_ & vbTab & k & "=" & CStr(v)
End Sub

' op_read_float to the end: the samples to <base>.f32; returns the samples per channel, the time, the holes
Private Function DecodeFloat(ByVal of_ As Long, ByVal outPath As String, ByVal channels As Long, ms_ As Double, holes As Long, errs As String) As Double
    Dim buf() As Single, h As Long, n As Long, li As Long, total As Double, t0 As Currency, t1 As Currency, w As Long
    ReDim buf(0 To 120& * 48 * 8 - 1)
    h = OutOpen(outPath)
    Do
        QueryPerformanceCounter t0
        n = op_read_float(of_, buf, 0, 120& * 48 * 8, li)
        QueryPerformanceCounter t1
        ms_ = ms_ + Ms(t0, t1)
        If n = OP_HOLE Then
            holes = holes + 1
        ElseIf n < 0 Then
            errs = errs & " " & n
            Exit Do
        ElseIf n = 0 Then
            Exit Do
        Else
            WriteFile h, buf(0), n * channels * 4, w, 0
            total = total + n
        End If
    Loop
    CloseHandle h
    DecodeFloat = total
End Function

Private Sub Main()
    Dim argv() As String, argc As Long, listPath As String, outDir As String, lines() As String, i As Long
    Dim f As Integer, res As String, noSeek As Boolean, text As String
    QueryPerformanceFrequency qpcFreq
    argc = GetArgs(argv)
    If argc < 3 Then Exit Sub
    listPath = argv(1)
    outDir = argv(2)
    If argc > 3 Then noSeek = (argv(3) = "noseek")
    logPath = outDir & "\log.txt"
    f = FreeFile
    Open listPath For Input As #f
    text = Input$(LOF(f), #f)
    Close #f
    lines = Split(Replace(text, vbCr, ""), vbLf)
    f = FreeFile
    Open outDir & "\results.txt" For Output As #f
    Close #f
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            res = TestFile(i, lines(i), outDir, noSeek)
            f = FreeFile
            Open outDir & "\results.txt" For Append As #f
            Print #f, res
            Close #f
        End If
    Next
    f = FreeFile
    Open outDir & "\results.txt" For Append As #f
    Print #f, "done" & vbTab & "log=" & logCount & vbTab & "asserts=" & assertCount
    Close #f
End Sub

Private Function TestFile(ByVal n As Long, ByVal path As String, ByVal outDir As String, ByVal noSeek As Boolean) As String
    Dim data() As Byte, size As Long, of_ As Long, err_ As Long, head As OpusHead, line_ As String, base As String
    Dim channels As Long, ms_ As Double, holes As Long, errs As String, total As Double, t0 As Currency, t1 As Currency
    Dim buf16() As Integer, buf2() As Single, h As Long, k As Long, w As Long, li As Long, nn As Long
    Dim dec As WZOpusDecoder, wzd As WZDecoder, pcm() As Byte, written As Long, wzTotal As Double, mode As Long
    Dim targets(0 To 4) As Double, ret As Long, pcm_total As Double, tags As OpusTags
    On Error GoTo fail
    base = outDir & "\" & n
    line_ = CStr(n)
    stage = "read"
    size = ReadBytes(path, data)

    ' op_open_memory, op_read_float
    stage = "open"
    QueryPerformanceCounter t0
    of_ = op_open_memory(data, size, err_)
    QueryPerformanceCounter t1
    If of_ = 0 Then
        AddField line_, "open_error", err_
        TestFile = line_
        Exit Function
    End If
    AddField line_, "open_ms", Format$(Ms(t0, t1), "0.000")
    head = op_head(of_, -1)
    channels = head.channel_count
    pcm_total = op_pcm_total(of_, -1)
    AddField line_, "channels", channels
    AddField line_, "pcm_total", pcm_total
    AddField line_, "links", op_link_count(of_)
    AddField line_, "pre_skip", head.pre_skip
    AddField line_, "gain", head.output_gain
    AddField line_, "mapping_family", head.mapping_family
    AddField line_, "streams", head.stream_count
    AddField line_, "coupled", head.coupled_count
    AddField line_, "serialno", op_serialno(of_, -1)
    AddField line_, "seekable", op_seekable(of_)
    AddField line_, "raw_total", op_raw_total(of_, -1)
    AddField line_, "bitrate", op_bitrate(of_, -1)
    If op_tags(of_, -1, tags) Then AddField line_, "comments", tags.comments
    stage = "float"
    ms_ = 0: holes = 0: errs = ""
    total = DecodeFloat(of_, base & ".f32", channels, ms_, holes, errs)
    AddField line_, "float_samples", total
    AddField line_, "float_ms", Format$(ms_, "0.0")
    AddField line_, "float_holes", holes
    If Len(errs) Then AddField line_, "float_errors", Trim$(errs)
    AddField line_, "pcm_tell_end", op_pcm_tell(of_)
    AddField line_, "bitrate_instant", op_bitrate_instant(of_)
    op_free of_

    ' op_read_stereo (16-bit, dithered)
    stage = "s16"
    of_ = op_open_memory(data, size, err_)
    ReDim buf16(0 To 11520 - 1)
    h = OutOpen(base & ".s16")
    ms_ = 0: holes = 0: errs = "": total = 0
    Do
        QueryPerformanceCounter t0
        nn = op_read_stereo(of_, buf16, 0, 11520)
        QueryPerformanceCounter t1
        ms_ = ms_ + Ms(t0, t1)
        If nn = OP_HOLE Then
            holes = holes + 1
        ElseIf nn < 0 Then
            errs = errs & " " & nn
            Exit Do
        ElseIf nn = 0 Then
            Exit Do
        Else
            WriteFile h, buf16(0), nn * 4, w, 0
            total = total + nn
        End If
    Loop
    CloseHandle h
    AddField line_, "s16_samples", total
    AddField line_, "s16_ms", Format$(ms_, "0.0")
    AddField line_, "s16_holes", holes
    If Len(errs) Then AddField line_, "s16_errors", Trim$(errs)
    op_free of_

    ' op_read_float_stereo (the downmix) for more than 2 channels
    If channels > 2 Then
        stage = "f2"
        of_ = op_open_memory(data, size, err_)
        ReDim buf2(0 To 11520 - 1)
        h = OutOpen(base & ".f2")
        total = 0
        Do
            nn = op_read_float_stereo(of_, buf2, 0, 11520)
            If nn = OP_HOLE Then
            ElseIf nn <= 0 Then
                Exit Do
            Else
                WriteFile h, buf2(0), nn * 8, w, 0
                total = total + nn
            End If
        Loop
        CloseHandle h
        AddField line_, "f2_samples", total
        op_free of_
    End If

    ' WZOpusDecoder, as openal_track's sound_PlayStream and sound_UpdateStream use it
    For mode = 0 To 1
        stage = "wz" & mode
        Set dec = WZOpusDecoder_fromFilename(path, mode = 1)
        If dec Is Nothing Then
            AddField line_, IIf(mode = 0, "wz", "wzm") & "_error", "nullptr"
        Else
            Set wzd = dec
            Set dec = Nothing
            If mode = 0 Then
                AddField line_, "wz_totalTime", wzd.totalTime()
                AddField line_, "wz_channels", wzd.channels()
                AddField line_, "wz_frequency", wzd.frequency()
            End If
            ReDim pcm(0 To 16384 - 1)
            h = OutOpen(base & IIf(mode = 0, ".wz", ".wzm"))
            wzTotal = 0: ms_ = 0
            Do
                QueryPerformanceCounter t0
                If Not wzd.decode(pcm, 16384, written) Then
                    AddField line_, IIf(mode = 0, "wz", "wzm") & "_nullopt", 1
                    Exit Do
                End If
                QueryPerformanceCounter t1
                ms_ = ms_ + Ms(t0, t1)
                If written <= 0 Then Exit Do
                WriteFile h, pcm(0), written, w, 0
                wzTotal = wzTotal + written
            Loop
            CloseHandle h
            Set wzd = Nothing
            AddField line_, IIf(mode = 0, "wz", "wzm") & "_bytes", wzTotal
            AddField line_, IIf(mode = 0, "wz", "wzm") & "_ms", Format$(ms_, "0.0")
        End If
    Next

    ' op_pcm_seek, then 4800 samples; op_raw_seek
    If Not noSeek And pcm_total > 10000 Then
        stage = "seek"
        of_ = op_open_memory(data, size, err_)
        targets(0) = Int(pcm_total * 0.5)
        targets(1) = Int(pcm_total * 0.25)
        targets(2) = Int(pcm_total * 0.75) + 7
        targets(3) = pcm_total - 6000
        targets(4) = 0
        ReDim buf2(0 To 4800& * channels - 1)
        For k = 0 To 4
            ret = op_pcm_seek(of_, targets(k))
            AddField line_, "seek" & k, targets(k) & "/" & ret & "/" & op_pcm_tell(of_)
            h = OutOpen(base & ".seek" & k & ".f32")
            total = 0
            Do While total < 4800
                nn = op_read_float(of_, buf2, 0, (4800 - total) * channels, li)
                If nn <= 0 Then Exit Do
                WriteFile h, buf2(0), nn * channels * 4, w, 0
                total = total + nn
            Loop
            CloseHandle h
        Next
        ret = op_raw_seek(of_, Int(op_raw_total(of_, -1) / 2))
        nn = op_read_float(of_, buf2, 0, 4800& * channels, li)
        AddField line_, "raw_seek", ret & "/" & nn & "/" & op_pcm_tell(of_)
        op_free of_
    End If
    TestFile = line_
    Exit Function
fail:
    TestFile = CStr(n) & vbTab & "error=" & Err.Number & " " & Err.Description & " at " & stage
End Function
