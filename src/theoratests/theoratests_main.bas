Attribute VB_Name = "theoratests_main"
' TheoraTests.exe (VB-only, not part of the game; cluster K): decodes Theora streams with WzSequence's libtheora
' port (src/sequence/theora_*.bas, compiled in) through the legacy theora.h API as lib/sequence/sequence.cpp
' does, for tools/theora_test.py, which compares the frames with ffmpeg's decoder and the official
' libtheoradec.dll.
'
'   TheoraTests <list file> <output folder> <pp level: a number, or "max">
'
' The list: one packet file per line (tools/theora_test.py: per packet an int32 byte count, int32 b_o_s, int32
' e_o_s, then the bytes; the three header packets first). For the n-th file (0-based) the frames go to
' <output folder>\<n>.yuv: per decoded packet the picture region (theora_info's frame_width x frame_height at
' offset_x, offset_y) of theora_decode_YUVout's planes, Y then U then V (4:2:0: the chroma region is half the
' size, rounded up), as ffmpeg's -pix_fmt yuv420p; and a line to <output folder>\results.txt:
'   n TAB frames TAB pp level TAB decode ms TAB last granulepos TAB its theora_granule_time TAB w x h info
' where decode ms is the time spent in theora_decode_packetin and theora_decode_YUVout. A failure is
' "n TAB error TAB <what>". The last line: "total TAB <frames> TAB <decode ms>".
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

Private Function ReadInt32(b() As Byte, ByVal pos As Long) As Long
    Dim v As Long
    CopyMemory v, b(pos), 4
    ReadInt32 = v
End Function

' the next packet of a packet file at pos: False at its end
Private Function NextPacket(b() As Byte, ByVal n As Long, pos As Long, op As ogg_packet, ByVal packetno As Long) As Boolean
    Dim bytes As Long
    If pos + 12 > n Then Exit Function
    bytes = ReadInt32(b, pos)
    op.b_o_s = ReadInt32(b, pos + 4)
    op.e_o_s = ReadInt32(b, pos + 8)
    pos = pos + 12
    op.bytes = bytes
    If bytes > 0 Then
        ReDim op.packet(0 To bytes - 1)
        CopyMemory op.packet(0), b(pos), bytes
    Else
        Erase op.packet
    End If
    op.granulepos = -1
    op.packetno = packetno
    pos = pos + bytes
    NextPacket = True
End Function

' the picture region of a plane, appended to out at k
Private Sub CropPlane(plane() As Byte, ByVal stride As Long, ByVal x0 As Long, ByVal y0 As Long, ByVal w As Long, _
    ByVal h As Long, out() As Byte, k As Long)
    Dim row As Long
    For row = 0 To h - 1
        CopyMemory out(k), plane((y0 + row) * stride + x0), w
        k = k + w
    Next
End Sub

Private Function DecodeOne(ByVal pktPath As String, ByVal outPath As String, ByVal ppArg As String, _
    frames As Long, level As Long, decodeMs As Double, lastGranule As Double, lastTime As Double, _
    infoText As String) As String
    Dim b() As Byte, n As Long, pos As Long
    Dim op As ogg_packet
    Dim ti As theora_info
    Dim tc As theora_comment
    Dim td As theora_state
    Dim yuv As yuv_buffer
    Dim r As Long, i As Long, packetno As Long
    Dim ppMax As Long, pp As Long
    Dim freq As Currency, t0 As Currency, t1 As Currency, total As Currency
    Dim f As Integer
    Dim out() As Byte, frameSize As Long, k As Long
    Dim cw As Long, ch As Long
    Dim headers As Long

    n = ReadBytes(pktPath, b)
    theora_info_init ti
    theora_comment_init tc
    pos = 0
    packetno = 0
    ' the three header packets, as sequence.cpp's seq_Play reads them
    For headers = 0 To 2
        If Not NextPacket(b, n, pos, op, packetno) Then DecodeOne = "missing header": Exit Function
        packetno = packetno + 1
        r = theora_decode_header(ti, tc, op)
        If r < 0 Then DecodeOne = "theora_decode_header " & r: Exit Function
    Next
    r = theora_decode_init(td, ti)
    If r <> 0 Then DecodeOne = "theora_decode_init " & r: Exit Function
    r = theora_control(td, TH_DECCTL_GET_PPLEVEL_MAX, ppMax, 4)
    If r <> 0 Then DecodeOne = "TH_DECCTL_GET_PPLEVEL_MAX " & r: Exit Function
    If ppArg = "max" Then pp = ppMax Else pp = CLng(ppArg)
    r = theora_control(td, TH_DECCTL_SET_PPLEVEL, pp, 4)
    If r <> 0 Then DecodeOne = "TH_DECCTL_SET_PPLEVEL " & r: Exit Function
    level = pp
    infoText = ti.width & "x" & ti.height & " frame " & ti.frame_width & "x" & ti.frame_height & " at " & _
        ti.offset_x & "," & ti.offset_y & " fps " & ti.fps_numerator & "/" & ti.fps_denominator & _
        " pf " & ti.pixelformat & " vendor " & tc.vendor
    cw = (ti.frame_width + 1) \ 2
    ch = (ti.frame_height + 1) \ 2
    frameSize = ti.frame_width * ti.frame_height + 2 * cw * ch
    ReDim out(0 To frameSize - 1)

    If Len(Dir$(outPath)) > 0 Then Kill outPath
    f = FreeFile
    Open outPath For Binary Access Write As #f
    QueryPerformanceFrequency freq
    frames = 0
    Do While NextPacket(b, n, pos, op, packetno)
        packetno = packetno + 1
        QueryPerformanceCounter t0
        r = theora_decode_packetin(td, op)
        If r < 0 Then
            Close #f
            DecodeOne = "theora_decode_packetin " & r & " at packet " & packetno
            Exit Function
        End If
        r = theora_decode_YUVout(td, yuv)
        QueryPerformanceCounter t1
        total = total + (t1 - t0)
        If r < 0 Then
            Close #f
            DecodeOne = "theora_decode_YUVout " & r
            Exit Function
        End If
        k = 0
        CropPlane yuv.y, yuv.y_stride, ti.offset_x, ti.offset_y, ti.frame_width, ti.frame_height, out, k
        CropPlane yuv.u, yuv.uv_stride, ti.offset_x \ 2, ti.offset_y \ 2, cw, ch, out, k
        CropPlane yuv.v, yuv.uv_stride, ti.offset_x \ 2, ti.offset_y \ 2, cw, ch, out, k
        Put #f, , out
        frames = frames + 1
        lastGranule = td.granulepos
        lastTime = theora_granule_time(td, td.granulepos)
    Loop
    Close #f
    decodeMs = total * 1000# / freq
    theora_clear td
    theora_comment_clear tc
    theora_info_clear ti
    DecodeOne = ""
End Function

Private Sub WriteText(ByVal path As String, ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open path For Append As #f
    Print #f, text
    Close #f
End Sub

Sub Main()
    Dim argv() As String, argc As Long
    Dim listText As String, lines() As String, i As Long, n As Long
    Dim f As Integer, outDir As String, results As String
    Dim frames As Long, level As Long, ms As Double, gp As Double, tm As Double, info As String
    Dim err_ As String, totalFrames As Long, totalMs As Double
    On Error GoTo fail
    argc = GetArgs(argv)
    If argc < 4 Then ExitProcess 2
    f = FreeFile
    Open argv(1) For Input As #f
    listText = Input$(LOF(f), f)
    Close #f
    lines = Split(Replace(listText, vbCr, ""), vbLf)
    outDir = argv(2)
    results = outDir & "\results.txt"
    If Len(Dir$(results)) > 0 Then Kill results
    n = 0
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            frames = 0
            ms = 0
            info = ""
            On Error Resume Next
            err_ = DecodeOne(lines(i), outDir & "\" & n & ".yuv", argv(3), frames, level, ms, gp, tm, info)
            If Err.Number <> 0 Then err_ = "VB error " & Err.Number & ": " & Err.Description
            On Error GoTo fail
            If Len(err_) > 0 Then
                WriteText results, n & vbTab & "error" & vbTab & err_
            Else
                WriteText results, n & vbTab & frames & vbTab & level & vbTab & Replace(Format$(ms, "0.000"), ",", ".") & vbTab & _
                    gp & vbTab & Replace(Format$(tm, "0.000000"), ",", ".") & vbTab & info
                totalFrames = totalFrames + frames
                totalMs = totalMs + ms
            End If
            n = n + 1
        End If
    Next
    WriteText results, "total" & vbTab & totalFrames & vbTab & Replace(Format$(totalMs, "0.000"), ",", ".")
    ExitProcess 0
fail:
    WriteText outDir & "\results.txt", "fatal" & vbTab & Err.Number & " " & Err.Description
    ExitProcess 1
End Sub
