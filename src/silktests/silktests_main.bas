Attribute VB_Name = "silktests_main"
' SilkTests.exe (VB-only, not part of the game; cluster K, fork O3): decodes Opus packet files with WzSound's
' libopus port (src/sound/silk_*, opus_*, and O2's entcode_c and celt_* compiled in) for
' src/silktests/silk_test.py, which compares the samples with ffmpeg's libopus decoder.
'
'   SilkTests <list file> <output folder>
'
' The list: one packet file per line (written by silk_test.py): int32 channels, int32 count, int32 Fs (the
' decoder's rate), int32 streams, int32 coupled streams, the channel mapping (channels bytes, padded to 4), then
' per packet int32 len, int32 decode_fec, int32 frame_size, len bytes (len 0: a lost packet, concealed by the PLC;
' decode_fec 1: the FEC of the next packet). For the n-th file (0-based): opus_decode_float's samples go to
' <output folder>\<n>.f32 (float little endian, interleaved), opus_decode's to <n>.s16 (with more than one
' stream: opus_multistream_decode_float's and opus_multistream_decode's), and a line to
' <output folder>\results.txt:
'   n TAB channels TAB samples TAB decode ms TAB multistream check TAB final range
' The multistream check decodes the packets again with opus_multistream_decode_float (one stream, the file's
' mapping) and says "same" when it gives opus_decode_float's samples; the final range is the XOR of
' OPUS_GET_FINAL_RANGE over the packets (hex). A failed decode is "n TAB error TAB <code or VB error>".
' The last line: "total TAB <decode ms of all files>".
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

Private Function I32At(b() As Byte, ByVal p As Long) As Long
    Dim v As Long
    CopyMemory v, b(p), 4
    I32At = v
End Function

Private Sub WriteFile_(ByVal path As String, ByVal addr As Long, ByVal nBytes As Long)
    Dim f As Integer, b() As Byte
    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    If nBytes > 0 Then
        ReDim b(0 To nBytes - 1)
        CopyMemory b(0), ByVal addr, nBytes
        Put #f, , b
    End If
    Close #f
End Sub

' decodes the packets of a packet file; api: 0 opus_decode_float, 1 opus_multistream_decode_float, 2 opus_decode
' (pcmS), 3 opus_multistream_decode (pcmS). Returns the sample count per channel (< 0: an error code), the final
' ranges' XOR in rng.
Private Function DecodeAll(data() As Byte, ByVal api As Long, pcmF() As Single, pcmS() As Integer, rng As Long) As Long
    Dim channels As Long, count As Long, p As Long, k As Long, len_ As Long, fec As Long, fsz As Long
    Dim st As Long, errc As Long, n As Long, total As Long, cap As Long, r As Long
    Dim Fs As Long, streams As Long, coupled As Long
    channels = I32At(data, 0)
    count = I32At(data, 4)
    Fs = I32At(data, 8)
    streams = I32At(data, 12)
    coupled = I32At(data, 16)
    p = 20 + ((channels + 3) \ 4) * 4                   ' the mapping, padded to 4 bytes
    cap = Fs * channels * 4
    ReDim pcmF(0 To cap - 1)
    ReDim pcmS(0 To cap - 1)
    If api = 1 Or api = 3 Then
        st = opus_multistream_decoder_create(Fs, channels, streams, coupled, data, 20, errc)
    Else
        st = opus_decoder_create(Fs, channels, errc)
    End If
    If st = 0 Then DecodeAll = errc: Exit Function
    rng = 0
    For k = 0 To count - 1
        len_ = I32At(data, p): fec = I32At(data, p + 4): fsz = I32At(data, p + 8)
        p = p + 12
        If (total + 5760) * channels > cap Then
            cap = cap * 2
            ReDim Preserve pcmF(0 To cap - 1)
            ReDim Preserve pcmS(0 To cap - 1)
        End If
        Select Case api
            Case 0: n = opus_decode_float(st, data, p, len_, pcmF, total * channels, fsz, fec)
            Case 1: n = opus_multistream_decode_float(st, data, p, len_, pcmF, total * channels, fsz, fec)
            Case 2: n = opus_decode(st, data, p, len_, pcmS, total * channels, fsz, fec)
            Case 3: n = opus_multistream_decode(st, data, p, len_, pcmS, total * channels, fsz, fec)
        End Select
        If n < 0 Then
            DecodeAll = n
            If api = 1 Or api = 3 Then opus_multistream_decoder_destroy st Else opus_decoder_destroy st
            Exit Function
        End If
        If api = 1 Or api = 3 Then
            opus_multistream_decoder_ctl st, OPUS_GET_FINAL_RANGE_REQUEST, r
        Else
            opus_decoder_ctl st, OPUS_GET_FINAL_RANGE_REQUEST, r
        End If
        rng = rng Xor r
        total = total + n
        p = p + len_
    Next
    If api = 1 Or api = 3 Then opus_multistream_decoder_destroy st Else opus_decoder_destroy st
    DecodeAll = total
End Function

Sub Main()
    Dim argv() As String, argc As Long, f As Integer, path As String, data() As Byte, idx As Long
    Dim channels As Long, samples As Long, out_ As String, res As Integer
    Dim freq As Currency, t0 As Currency, t1 As Currency, totalTime As Double, ms As Double
    Dim check As String, errNumber As Long, errText As String
    Dim pcmF() As Single, pcmS() As Integer, pcmF2() As Single, pcmS2() As Integer
    Dim rng As Long, rng2 As Long, samples2 As Long, i As Long, multi As Boolean
    argc = GetArgs(argv)
    If argc < 3 Then ExitProcess 2
    QueryPerformanceFrequency freq
    out_ = argv(2)
    res = FreeFile
    Open out_ & "\results.txt" For Output As #res
    f = FreeFile
    Open argv(1) For Input As #f
    idx = 0
    Do While Not EOF(f)
        Line Input #f, path
        path = Trim$(path)
        If Len(path) > 0 Then
            On Error GoTo failed
            ReadBytes path, data
            channels = I32At(data, 0)
            multi = (I32At(data, 12) > 1)               ' more than one stream: the multistream decoder only
            QueryPerformanceCounter t0
            If multi Then
                samples = DecodeAll(data, 1, pcmF, pcmS, rng)
            Else
                samples = DecodeAll(data, 0, pcmF, pcmS, rng)
            End If
            QueryPerformanceCounter t1
            ms = (t1 - t0) * 1000# / freq
            If samples < 0 Then
                Print #res, idx & vbTab & "error" & vbTab & samples & " " & opus_strerror(samples)
            Else
                totalTime = totalTime + ms
                WriteFile_ out_ & "\" & idx & ".f32", VarPtr(pcmF(0)), samples * channels * 4
                ' the multistream decoder: the same samples
                If multi Then
                    check = "same"
                Else
                    samples2 = DecodeAll(data, 1, pcmF2, pcmS2, rng2)
                    check = "same"
                    If samples2 <> samples Then
                        check = "samples " & samples2
                    ElseIf rng2 <> rng Then
                        check = "range " & Hex$(rng2)
                    Else
                        For i = 0 To samples * channels - 1
                            If pcmF2(i) <> pcmF(i) Then check = "sample " & i & ": " & pcmF2(i) & " / " & pcmF(i): Exit For
                        Next
                    End If
                End If
                ' opus_decode (16-bit, soft clipped)
                If multi Then
                    samples2 = DecodeAll(data, 3, pcmF2, pcmS2, rng2)
                Else
                    samples2 = DecodeAll(data, 2, pcmF2, pcmS2, rng2)
                End If
                If samples2 <> samples Then check = check & ", opus_decode samples " & samples2
                If samples2 > 0 Then WriteFile_ out_ & "\" & idx & ".s16", VarPtr(pcmS2(0)), samples2 * channels * 2
                Print #res, idx & vbTab & channels & vbTab & samples & vbTab & Format$(ms, "0.000") & vbTab & check & vbTab & Hex$(rng)
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
