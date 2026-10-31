Attribute VB_Name = "celttests_main"
' CeltTests.exe (VB-only, not part of the game; cluster K, fork O2): decodes CELT-only Opus packets with the
' port of libopus 1.5.2's CELT decoder and range decoder (src/sound/entcode_c.bas, celt_*.bas, compiled in)
' for tools/celt_test.py, which compares the samples with ffmpeg's libopus decoder. Each packet is decoded
' the way libopus' opus_decode_native and opus_decode_frame do for a CELT-only packet (src/opus_decoder.c):
' the TOC (stereo flag, bandwidth -> CELT_SET_END_BAND, frame size), opus_packet_parse_impl's frames of a code
' 0-3 packet, then per frame CELT_SET_END_BAND, CELT_SET_CHANNELS, CELT_SET_START_BAND(0), ec_dec_init and
' celt_decode_with_ec on the frame's bytes; a frame of 0 or 1 byte is a lost frame (data NULL: the PLC, with
' the last packet's frame size). The decoder is set up as opus_decoder_init does (celt_decoder_init at
' 48 kHz, CELT_SET_SIGNALLING(0)).
'
'   CeltTests <list file> <output folder>
'
' The list: one line per stream, "<packet file> TAB <name>". A packet file (tools/celt_test.py): "CELTPKT1",
' then int32 channels (the decoder's), int32 pre-skip, int32 packet count, int32 sampling rate (the decoder's),
' then per packet int32 length and its bytes.
' Output: <output folder>\<name>.f32 (the decoded samples, float, interleaved, from the first one: the
' pre-skip is the caller's), and a line per stream in <output folder>\results.txt:
'   name TAB channels TAB packets TAB samples TAB decode ms TAB lost frames TAB errors TAB final range TAB
'   error text TAB the decoder's path counters (celt_decoder_c's celt_vbCounts, comma-separated)
' or "name TAB error TAB <text>". The last line: "total TAB <decode ms of all streams>".
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

Private Sub WriteFloats(ByVal path As String, pcm() As Single, ByVal count As Long)
    Dim f As Integer, b() As Byte
    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    If count > 0 Then
        ReDim b(0 To count * 4 - 1)
        CopyMemory b(0), pcm(0), count * 4
        Put #f, , b
    End If
    Close #f
End Sub

Private Function I32(b() As Byte, ByVal p As Long) As Long
    Dim v As Long
    CopyMemory v, b(p), 4
    I32 = v
End Function

' opus_decoder.c's parse_size: the frame length's bytes (1 or 2), size set; -1 if the packet ends first
Private Function parse_size(b() As Byte, ByVal p As Long, ByVal len_ As Long, size As Long) As Long
    If len_ < 1 Then
        size = -1
        parse_size = -1
    ElseIf b(p) < 252 Then
        size = b(p)
        parse_size = 1
    ElseIf len_ < 2 Then
        size = -1
        parse_size = -1
    Else
        size = 4 * CLng(b(p + 1)) + b(p)
        parse_size = 2
    End If
End Function

' opus.c's opus_packet_parse_impl (self_delimited 0): the frames' offsets and sizes; the count or an error < 0
Private Function packet_parse(b() As Byte, ByVal p As Long, ByVal len_ As Long, frame_off() As Long, frame_len() As Long) As Long
    Dim toc As Long, count As Long, last_size As Long, bytes As Long, ch As Long, pad As Long, tmp As Long
    Dim i As Long, framesize As Long, q As Long, size0 As Long
    If len_ < 1 Then packet_parse = OPUS_INVALID_PACKET: Exit Function
    toc = b(p)
    framesize = 120 * (2 ^ ((toc \ 8) And 3))           ' opus_packet_get_samples_per_frame (CELT-only)
    q = p + 1
    len_ = len_ - 1
    last_size = len_
    Select Case toc And 3
    Case 0
        count = 1
    Case 1
        count = 2
        If (len_ And 1) <> 0 Then packet_parse = OPUS_INVALID_PACKET: Exit Function
        last_size = len_ \ 2
        frame_len(0) = last_size
    Case 2
        count = 2
        bytes = parse_size(b, q, len_, size0)
        len_ = len_ - bytes
        If size0 < 0 Or size0 > len_ Then packet_parse = OPUS_INVALID_PACKET: Exit Function
        q = q + bytes
        frame_len(0) = size0
        last_size = len_ - size0
    Case 3
        If len_ < 1 Then packet_parse = OPUS_INVALID_PACKET: Exit Function
        ch = b(q): q = q + 1
        count = ch And &H3F
        len_ = len_ - 1
        If count <= 0 Or framesize * count > 5760 Then packet_parse = OPUS_INVALID_PACKET: Exit Function
        If (ch And &H40) <> 0 Then
            Do
                If len_ <= 0 Then packet_parse = OPUS_INVALID_PACKET: Exit Function
                tmp = b(q): q = q + 1
                len_ = len_ - 1
                If tmp = 255 Then pad = 254 Else pad = tmp
                len_ = len_ - pad
            Loop While tmp = 255
        End If
        If len_ < 0 Then packet_parse = OPUS_INVALID_PACKET: Exit Function
        If (ch And &H80) <> 0 Then
            ' VBR case
            last_size = len_
            For i = 0 To count - 2
                bytes = parse_size(b, q, len_, size0)
                len_ = len_ - bytes
                If size0 < 0 Or size0 > len_ Then packet_parse = OPUS_INVALID_PACKET: Exit Function
                q = q + bytes
                frame_len(i) = size0
                last_size = last_size - bytes - size0
            Next
            If last_size < 0 Then packet_parse = OPUS_INVALID_PACKET: Exit Function
        Else
            ' CBR case
            last_size = len_ \ count
            If last_size * count <> len_ Then packet_parse = OPUS_INVALID_PACKET: Exit Function
            For i = 0 To count - 2
                frame_len(i) = last_size
            Next
        End If
    End Select
    If last_size > 1275 Then packet_parse = OPUS_INVALID_PACKET: Exit Function
    frame_len(count - 1) = last_size
    For i = 0 To count - 1
        frame_off(i) = q
        q = q + frame_len(i)
    Next
    packet_parse = count
End Function

' one stream: the packets of a packet file decoded into pcm(); the samples per channel
Private Function DecodeStream(data() As Byte, pcm() As Single, ByVal channels As Long, ByVal npackets As Long, ByVal rate As Long, _
        ms As Double, lost As Long, errors As Long, finalRange As Long, errText As String) As Long
    Dim st As Long, ret As Long, i As Long, k As Long, p As Long, plen As Long, toc As Long, count As Long
    Dim frame_off(0 To 47) As Long, frame_len(0 To 47) As Long, packet_frame_size As Long, endband As Long
    Dim stream_channels As Long, nb_samples As Long, dec As ec_ctx, t0 As Currency, t1 As Currency, freq As Currency
    Dim last_frame_size As Long, v As Long

    QueryPerformanceFrequency freq
    st = celt_decoder_vbNew()
    ret = celt_decoder_init(st, rate, channels)
    If ret <> OPUS_OK Then errText = "celt_decoder_init " & ret: Exit Function
    celt_decoder_ctl st, CELT_SET_SIGNALLING_REQUEST, 0
    QueryPerformanceCounter t0

    p = 24
    nb_samples = 0
    For i = 0 To npackets - 1
        plen = I32(data, p)
        p = p + 4
        toc = data(p)
        If (toc And &H80) = 0 Then errText = "packet " & i & " is not CELT-only (TOC " & toc & ")": Exit Function
        ' opus_packet_get_bandwidth for a CELT-only TOC: MB (0) is NB
        Select Case (toc \ 32) And 3
        Case 0: endband = 13                            ' OPUS_BANDWIDTH_NARROWBAND
        Case 1: endband = 17                            ' OPUS_BANDWIDTH_WIDEBAND
        Case 2: endband = 19                            ' OPUS_BANDWIDTH_SUPERWIDEBAND
        Case 3: endband = 21                            ' OPUS_BANDWIDTH_FULLBAND
        End Select
        packet_frame_size = (120 * (2 ^ ((toc \ 8) And 3))) * rate \ 48000   ' opus_packet_get_samples_per_frame(data, Fs)
        If (toc And 4) <> 0 Then stream_channels = 2 Else stream_channels = 1
        count = packet_parse(data, p, plen, frame_off, frame_len)
        If count < 0 Then errText = "packet " & i & ": parse error " & count: Exit Function
        If UBound(pcm) < (nb_samples + count * packet_frame_size) * channels - 1 Then
            ReDim Preserve pcm(0 To (nb_samples + count * packet_frame_size) * channels * 2)
        End If
        For k = 0 To count - 1
            If frame_len(k) <= 1 Then
                ' opus_decode_frame: data = NULL; frame_size = IMIN(frame_size, st->frame_size) (the PLC runs with
                ' st->stream_channels, set from this packet's TOC by opus_decode_native; no CELT_SET_END_BAND:
                ' bandwidth is 0)
                celt_decoder_ctl st, CELT_SET_CHANNELS_REQUEST, stream_channels
                celt_decoder_ctl st, CELT_SET_START_BAND_REQUEST, 0
                ret = celt_decode_with_ec(st, data, 0, True, 0, pcm, nb_samples * channels, packet_frame_size, dec, False, 0)
                lost = lost + 1
            Else
                ec_dec_init dec, data, frame_off(k), frame_len(k)
                celt_decoder_ctl st, CELT_SET_END_BAND_REQUEST, endband
                celt_decoder_ctl st, CELT_SET_CHANNELS_REQUEST, stream_channels
                celt_decoder_ctl st, CELT_SET_START_BAND_REQUEST, 0
                ret = celt_decode_with_ec(st, data, frame_off(k), False, frame_len(k), pcm, nb_samples * channels, packet_frame_size, dec, False, 0)
            End If
            If ret < 0 Then
                errors = errors + 1
                errText = "packet " & i & " frame " & k & ": celt_decode_with_ec " & ret
                ret = packet_frame_size
            End If
            If ret <> packet_frame_size Then errText = "packet " & i & ": " & ret & " samples, expected " & packet_frame_size
            celt_decoder_ctl st, CELT_GET_AND_CLEAR_ERROR_REQUEST, v
            If v <> 0 Then errors = errors + 1
            nb_samples = nb_samples + ret
        Next
        p = p + plen
    Next
    QueryPerformanceCounter t1
    ms = (t1 - t0) / freq * 1000#
    celt_decoder_ctl st, OPUS_GET_FINAL_RANGE_REQUEST, finalRange
    celt_decoder_vbFree st
    DecodeStream = nb_samples
End Function

Sub Main()
    Dim argv() As String, argc As Long, lines() As String, i As Long, parts() As String, outDir As String
    Dim fres As Integer, data() As Byte, n As Long, pcm() As Single, samples As Long, ms As Double, total As Double
    Dim channels As Long, npackets As Long, lost As Long, errors As Long, finalRange As Long, errText As String
    Dim listText As String, f As Integer, rate As Long, counts As String, k As Long

    On Error GoTo fail
    argc = GetArgs(argv)
    If argc < 3 Then ExitProcess 2
    outDir = argv(2)
    f = FreeFile
    Open argv(1) For Binary Access Read As #f
    listText = Space$(LOF(f))
    Get #f, , listText
    Close #f
    lines = Split(Replace(listText, vbCr, ""), vbLf)

    fres = FreeFile
    Open outDir & "\results.txt" For Output As #fres
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            parts = Split(lines(i), vbTab)
            n = ReadBytes(parts(0), data)
            If n < 24 Then
                Print #fres, parts(1) & vbTab & "error" & vbTab & "short packet file"
            Else
                channels = I32(data, 8)
                npackets = I32(data, 16)
                rate = I32(data, 20)
                For k = 0 To 15
                    celt_vbCounts(k) = 0
                Next
                ReDim pcm(0 To 48000 * 2 * channels)
                lost = 0: errors = 0: errText = ""
                samples = 0
                On Error Resume Next
                samples = DecodeStream(data, pcm, channels, npackets, rate, ms, lost, errors, finalRange, errText)
                If Err.Number <> 0 Then errText = "VB error " & Err.Number & ": " & Err.Description
                On Error GoTo fail
                If Len(errText) > 0 And samples = 0 Then
                    Print #fres, parts(1) & vbTab & "error" & vbTab & errText
                ElseIf samples < 0 Or samples * channels > UBound(pcm) + 1 Then
                    Print #fres, parts(1) & vbTab & "error" & vbTab & samples & " samples (" & errText & ")"
                Else
                    WriteFloats outDir & "\" & parts(1) & ".f32", pcm, samples * channels
                    counts = CStr(celt_vbCounts(0))
                    For k = 1 To 15
                        counts = counts & "," & celt_vbCounts(k)
                    Next
                    Print #fres, parts(1) & vbTab & channels & vbTab & npackets & vbTab & samples & vbTab & Format$(ms, "0.000") & vbTab & lost & vbTab & errors & vbTab & Hex$(finalRange) & vbTab & errText & vbTab & counts
                    total = total + ms
                End If
            End If
        End If
    Next
    Print #fres, "total" & vbTab & Format$(total, "0.000")
    Close #fres
    ExitProcess 0
fail:
    errText = Err.Number & " " & Err.Description & " (stream " & i & ")"
    Close
    Resume fatal
fatal:
    On Error Resume Next
    fres = FreeFile
    Open outDir & "\results.txt" For Append As #fres
    Print #fres, "fatal" & vbTab & "error" & vbTab & errText
    Close #fres
    ExitProcess 1
End Sub
