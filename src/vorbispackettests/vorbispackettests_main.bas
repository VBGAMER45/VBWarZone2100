Attribute VB_Name = "vorbispackettests_main"
' VorbisPacketTests.exe (VB-only, not part of the game; cluster K): runs the videos' Vorbis packets through
' WzSound's libvorbis-style synthesis API (src/sound/stb_vorbis_c.bas, compiled in) as lib/sequence's
' sequence.cpp does, for tools/video_audio_test.py, which compares the samples with ffmpeg's decoder.
'
'   VorbisPacketTests <list file> <output folder>
'
' The list: one packet file per line (video_audio_test.py's write_packet_file: an int32 count, then per packet
' bytes, b_o_s, e_o_s as int32, granulepos, packetno as float64, then the bytes). For the n-th file (0-based)
' the samples (float32, interleaved) go to <output folder>\<n>.f32, and a line to results.txt:
'   n TAB channels TAB rate TAB samples TAB decode ms TAB checks
' checks is "ok" or the checks that failed, of: the three header packets accepted; libvorbis's refusals
' (a comment or setup header before its turn, an identification header twice or without b_o_s, a packet
' that is not Vorbis, a dsp state on an incomplete vorbis_info, a header packet given to vorbis_synthesis,
' a block while samples are pending, a read past the samples); every audio packet decoded; after each packet
' with a granule position, vd.granulepos equal to it and the samples returned so far too; the samples taken
' in odd-sized pieces (pcmout, then read of part of them). A failed file is "n TAB error TAB <what>". The last
' line: "total TAB <decode ms of all files>".
Option Explicit

Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' codec.h
Private Const OV_EINVAL As Long = -131
Private Const OV_ENOTVORBIS As Long = -132
Private Const OV_EBADHEADER As Long = -133
Private Const OV_ENOTAUDIO As Long = -135

' a packet file's packets: their fields and where their bytes are in the file
Private pkCount As Long
Private pkBytes() As Long
Private pkBos() As Long
Private pkEos() As Long
Private pkGranule() As Double
Private pkNo() As Double
Private pkOffset() As Long
Private fileData() As Byte

Private failures As String

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

Private Sub WriteFloats(ByVal path As String, v() As Single, ByVal count As Long)
    Dim f As Integer, b() As Byte
    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    If count > 0 Then
        ReDim b(0 To count * 4 - 1)
        CopyMemory b(0), v(0), count * 4
        Put #f, , b
    End If
    Close #f
End Sub

' the packet file's index
Private Sub LoadPackets(ByVal path As String)
    Dim n As Long, p As Long, i As Long
    n = ReadBytes(path, fileData)
    CopyMemory pkCount, fileData(0), 4
    ReDim pkBytes(0 To pkCount), pkBos(0 To pkCount), pkEos(0 To pkCount)
    ReDim pkGranule(0 To pkCount), pkNo(0 To pkCount), pkOffset(0 To pkCount)
    p = 4
    For i = 0 To pkCount - 1
        CopyMemory pkBytes(i), fileData(p), 4
        CopyMemory pkBos(i), fileData(p + 4), 4
        CopyMemory pkEos(i), fileData(p + 8), 4
        CopyMemory pkGranule(i), fileData(p + 12), 8
        CopyMemory pkNo(i), fileData(p + 20), 8
        pkOffset(i) = p + 28
        p = p + 28 + pkBytes(i)
    Next
End Sub

' packet i's bytes from index 0 (the ogg_packet's packet)
Private Sub PacketBytes(ByVal i As Long, pkt() As Byte)
    If pkBytes(i) > 0 Then
        ReDim pkt(0 To pkBytes(i) - 1)
        CopyMemory pkt(0), fileData(pkOffset(i)), pkBytes(i)
    Else
        Erase pkt
    End If
End Sub

Private Sub Fail(ByVal what As String)
    If Len(failures) > 400 Then Exit Sub
    If InStr(failures, what) = 0 Then
        If Len(failures) > 0 Then failures = failures & "; "
        failures = failures & what
    End If
End Sub

Private Sub Expect(ByVal got As Long, ByVal want As Long, ByVal what As String)
    If got <> want Then Fail what & " " & got & " (expected " & want & ")"
End Sub

' one video's packets through the API; the samples to out (interleaved), their count returned
Private Function DecodeFile(outv() As Single, channels As Long, rate As Long) As Long
    Dim vi As Long, vc As Long, vd As Long, vb As Long, vi2 As Long, vc2 As Long, vd2 As Long
    Dim pkt() As Byte, fake(0 To 29) As Byte, i As Long, r As Long, n As Long, k As Long, j As Long, s As Long
    Dim pcm() As Single, total As Long, cap As Long, chunk As Long, iter As Long, readChecked As Boolean
    Dim audio As Long, failedAudio As Long

    ' the headers
    vorbis_info_init vi
    vorbis_comment_init vc
    For i = 0 To 2
        PacketBytes i, pkt
        r = vorbis_synthesis_headerin(vi, vc, pkt, pkBytes(i), pkBos(i), pkEos(i), pkGranule(i), pkNo(i))
        Expect r, 0, "headerin(" & i & ")"
    Next
    channels = vorbis_info_channels(vi)
    rate = vorbis_info_rate(vi)

    ' libvorbis's refusals
    vorbis_info_init vi2
    vorbis_comment_init vc2
    PacketBytes 1, pkt
    Expect vorbis_synthesis_headerin(vi2, vc2, pkt, pkBytes(1), 0, 0, -1, 1), OV_EBADHEADER, "comment before info"
    PacketBytes 0, pkt
    Expect vorbis_synthesis_headerin(vi2, vc2, pkt, pkBytes(0), 0, 0, 0, 0), OV_EBADHEADER, "info without b_o_s"
    fake(0) = &H80: fake(1) = &H74: fake(2) = &H68: fake(3) = &H65: fake(4) = &H6F: fake(5) = &H72: fake(6) = &H61
    Expect vorbis_synthesis_headerin(vi2, vc2, fake, 30, 1, 0, 0, 0), OV_ENOTVORBIS, "theora header"
    Expect vorbis_synthesis_headerin(vi2, vc2, pkt, pkBytes(0), 1, 0, 0, 0), 0, "info"
    Expect vorbis_synthesis_headerin(vi2, vc2, pkt, pkBytes(0), 1, 0, 0, 0), OV_EBADHEADER, "info twice"
    PacketBytes 2, pkt
    Expect vorbis_synthesis_headerin(vi2, vc2, pkt, pkBytes(2), 0, 0, 0, 2), OV_EBADHEADER, "setup before comment"
    Expect vorbis_synthesis_init(vd2, vi2), 1, "synthesis_init without setup"
    Expect vd2, 0, "the failed dsp state's handle"
    If pkCount > 3 Then
        PacketBytes 3, pkt
        Expect vorbis_synthesis_headerin(vi, vc, pkt, pkBytes(3), pkBos(3), pkEos(3), pkGranule(3), pkNo(3)), OV_ENOTVORBIS, "audio packet as a header"
    End If
    vorbis_comment_clear vc2
    vorbis_info_clear vi2
    Expect vi2, 0, "the cleared info's handle"

    ' the decoder
    Expect vorbis_synthesis_init(vd, vi), 0, "synthesis_init"
    Expect vorbis_block_init(vd, vb), 0, "block_init"
    PacketBytes 0, pkt
    Expect vorbis_synthesis(vb, pkt, pkBytes(0), 1, 0, 0, 0), OV_ENOTAUDIO, "header packet to vorbis_synthesis"

    cap = 65536
    ReDim outv(0 To cap - 1)
    total = 0
    For i = 3 To pkCount - 1
        PacketBytes i, pkt
        audio = audio + 1
        r = vorbis_synthesis(vb, pkt, pkBytes(i), pkBos(i), pkEos(i), pkGranule(i), pkNo(i))
        If r = 0 Then
            Expect vorbis_synthesis_blockin(vd, vb), 0, "blockin"
            If vorbis_synthesis_pcmout(vd, pcm) > 0 Then
                Expect vorbis_synthesis_blockin(vd, vb), OV_EINVAL, "blockin while samples are pending"
            End If
            If pkGranule(i) <> -1 Then
                If vorbis_dsp_granulepos(vd) <> pkGranule(i) Then Fail "vd.granulepos " & vorbis_dsp_granulepos(vd) & " at packet " & i & " (" & pkGranule(i) & ")"
            End If
            Do
                n = vorbis_synthesis_pcmout(vd, pcm)
                If n = 0 Then Exit Do
                If Not readChecked Then
                    Expect vorbis_synthesis_read(vd, n + 1), OV_EINVAL, "read past the samples"
                    readChecked = True
                End If
                iter = iter + 1
                chunk = 1 + (iter * 97) Mod 509
                If chunk > n Then k = n Else k = chunk
                If (total + k) * channels > cap Then
                    Do While (total + k) * channels > cap
                        cap = cap * 2
                    Loop
                    ReDim Preserve outv(0 To cap - 1)
                End If
                s = total * channels
                For j = 0 To k - 1
                    For r = 0 To channels - 1
                        outv(s) = pcm(r, j)
                        s = s + 1
                    Next
                Next
                total = total + k
                Expect vorbis_synthesis_read(vd, k), 0, "read"
            Loop
            If pkGranule(i) <> -1 Then
                If total <> pkGranule(i) Then Fail "samples " & total & " at packet " & i & " (granule position " & pkGranule(i) & ")"
            End If
        Else
            failedAudio = failedAudio + 1
            Fail "synthesis " & r & " at packet " & i
        End If
    Next

    Expect vorbis_block_clear(vb), 0, "block_clear"
    vorbis_dsp_clear vd
    vorbis_comment_clear vc
    vorbis_info_clear vi
    Expect vb + vd + vc + vi, 0, "the cleared handles"
    DecodeFile = total
End Function

Sub Main()
    Dim argv() As String, argc As Long, f As Integer, path As String, idx As Long, res As Integer
    Dim channels As Long, rate As Long, outv() As Single, samples As Long, out As String
    Dim freq As Currency, t0 As Currency, t1 As Currency, totalTime As Double, ms As Double
    Dim errNumber As Long, errText As String
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
            LoadPackets path
            failures = ""
            QueryPerformanceCounter t0
            samples = DecodeFile(outv, channels, rate)
            QueryPerformanceCounter t1
            ms = (t1 - t0) * 1000# / freq
            totalTime = totalTime + ms
            WriteFloats out & "\" & idx & ".f32", outv, samples * channels
            If Len(failures) = 0 Then failures = "ok"
            Print #res, idx & vbTab & channels & vbTab & rate & vbTab & samples & vbTab & Format$(ms, "0.000") & vbTab & failures
            GoTo next_
failed:
            errNumber = Err.Number: errText = Err.Description
            Print #res, idx & vbTab & "error" & vbTab & "VB error " & errNumber & ": " & errText & " (" & failures & ")"
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
