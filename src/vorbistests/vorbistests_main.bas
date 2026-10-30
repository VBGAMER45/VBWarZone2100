Attribute VB_Name = "vorbistests_main"
' VorbisTests.exe (VB-only, not part of the game; cluster H): decodes Ogg Vorbis files with WzSound's
' stb_vorbis port (src/sound/stb_vorbis_c.bas, compiled in) for tools/run_tests.py --sound, which compares
' the samples with ffmpeg's decoder.
'
'   VorbisTests <list file> <output folder>
'
' The list: one file path per line. For the n-th file (0-based): stb_vorbis_decode_memory's samples go to
' <output folder>\<n>.pcm (16-bit little endian, interleaved), and a line to <output folder>\results.txt:
'   n TAB channels TAB rate TAB samples TAB decode ms TAB stream_length_in_samples TAB stream check
' The stream check decodes the file again as WzSound's decoder does (stb_vorbis_open_memory,
' stb_vorbis_get_samples_short_interleaved in 4096-short calls, after a stb_vorbis_seek_start and a first
' read) and says "same" when it gives decode_memory's samples, else what differs. A failed decode is
' "n TAB error TAB <code or VB error>". The last line: "total TAB <decode ms of all files>".
Option Explicit

Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

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

Private Sub WriteSamples(ByVal path As String, pcm() As Integer, ByVal count As Long)
    Dim f As Integer, b() As Byte
    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    If count > 0 Then
        ReDim b(0 To count * 2 - 1)
        CopyMemory b(0), pcm(0), count * 2
        Put #f, , b
    End If
    Close #f
End Sub

' the stream check: the file decoded again through the streaming calls, against decode_memory's samples
Private Function StreamCheck(data() As Byte, ByVal n As Long, pcm() As Integer, ByVal samples As Long, ByVal channels As Long, streamLen As Long) As String
    Dim v As Long, e As Long, info As stb_vorbis_info, buf() As Integer, got As Long, total As Long, i As Long
    Dim first As Long, firstCopy() As Integer
    v = stb_vorbis_open_memory(data, n, e)
    If v = 0 Then StreamCheck = "open error " & e: Exit Function
    info = stb_vorbis_get_info(v)
    If info.channels <> channels Then StreamCheck = "channels " & info.channels: stb_vorbis_close v: Exit Function
    streamLen = stb_vorbis_stream_length_in_samples(v)
    ' a first read, then back to the start (as a looping stream does)
    ReDim buf(0 To 4095)
    first = stb_vorbis_get_samples_short_interleaved(v, channels, buf, 0, 4096)
    If stb_vorbis_seek_start(v) = 0 And samples > 0 Then StreamCheck = "seek_start failed": stb_vorbis_close v: Exit Function
    ReDim buf(0 To samples * channels + 4096)
    total = 0
    Do
        got = stb_vorbis_get_samples_short_interleaved(v, channels, buf, total * channels, 4096)
        If got = 0 Then Exit Do
        total = total + got
        If (total + 4096) * channels > UBound(buf) Then ReDim Preserve buf(0 To (total + 4096) * channels * 2)
    Loop
    stb_vorbis_close v
    If total <> samples Then StreamCheck = "samples " & total: Exit Function
    For i = 0 To samples * channels - 1
        If buf(i) <> pcm(i) Then StreamCheck = "sample " & i & ": " & buf(i) & " / " & pcm(i): Exit Function
    Next
    StreamCheck = "same"
End Function

Sub Main()
    Dim argv() As String, argc As Long, f As Integer, path As String, data() As Byte, n As Long, idx As Long
    Dim channels As Long, rate As Long, pcm() As Integer, samples As Long, out As String, res As Integer
    Dim freq As Currency, t0 As Currency, t1 As Currency, totalTime As Double, ms As Double, streamLen As Long
    Dim check As String, errNumber As Long, errText As String
    argc = GetArgs(argv)
    If argc < 3 Then ExitProcess 2
    QueryPerformanceFrequency freq
    out = argv(2)
    res = FreeFile
    Open out & "\results.txt" For Output As #res
    f = FreeFile
    Open argv(1) For Input As #f
    idx = 0
    Do While Not EOF(f)
        Line Input #f, path
        path = Trim$(path)
        If Len(path) > 0 Then
            On Error GoTo failed
            n = ReadBytes(path, data)
            QueryPerformanceCounter t0
            samples = stb_vorbis_decode_memory(data, n, channels, rate, pcm)
            QueryPerformanceCounter t1
            ms = (t1 - t0) * 1000# / freq
            If samples < 0 Then
                Print #res, idx & vbTab & "error" & vbTab & samples
            Else
                totalTime = totalTime + ms
                WriteSamples out & "\" & idx & ".pcm", pcm, samples * channels
                streamLen = 0
                check = StreamCheck(data, n, pcm, samples, channels, streamLen)
                Print #res, idx & vbTab & channels & vbTab & rate & vbTab & samples & vbTab & Format$(ms, "0.000") & vbTab & streamLen & vbTab & check
            End If
            GoTo next_
failed:
            errNumber = Err.Number: errText = Err.Description
            Print #res, idx & vbTab & "error" & vbTab & "VB error " & errNumber & ": " & errText
            Resume next_
next_:
            On Error GoTo 0
            idx = idx + 1
        End If
    Loop
    Close #f
    Print #res, "total" & vbTab & Format$(totalTime, "0.0")
    Close #res
    ExitProcess 0
End Sub
