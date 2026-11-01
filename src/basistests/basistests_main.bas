Attribute VB_Name = "basistests_main"
' BasisTests.exe (VB-only, not part of the game; cluster T, fork T2): transcodes KTX2 files with the port of Basis
' Universal's transcoder (src/basis's basisu_* modules and ktx2_transcoder, compiled in, with the zstd decoder's
' modules) for tools/ktx2_test.py, which compares the output with tools/ktx2ref.cpp's (the C, as 4.7.0 builds it).
' Each file: basisu_transcoder_init (once), ktx2_transcoder.init, start_transcoding, then every level (layer 0, face
' 0) transcoded with decode_flags 0 as gfx_api_image_basis_priv.cpp does it (RGBA32 etc. at the level's original
' size, block formats as m_total_blocks blocks), the levels' outputs one after another.
'
'   BasisTests <format> <output folder> <list file> [nowrite]
'
' format: a transcoder_texture_format number; the list: one KTX2 file path per line.
' Output: <output folder>\<file name>.<format> (unless nowrite), and a line per file in <output folder>\results.txt:
'   name TAB width TAB height TAB levels TAB bytes TAB fnv (FNV-1a 32 of the output, %08x) TAB ms (init + transcode)
' or "name TAB error TAB <text>". The last line: "total TAB <ms of all files> TAB <init ms>".
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

Private Sub WriteBytes(ByVal path As String, b() As Byte, ByVal n As Long)
    Dim f As Integer, c() As Byte
    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    If n > 0 Then
        If n = UBound(b) + 1 Then
            Put #f, , b
        Else
            ReDim c(0 To n - 1)
            CopyMemory c(0), b(0), n
            Put #f, , c
        End If
    End If
    Close #f
End Sub

Private Function Now_() As Double
    Static freq As Currency
    Dim c As Currency
    If freq = 0 Then QueryPerformanceFrequency freq
    QueryPerformanceCounter c
    Now_ = CDbl(c) / CDbl(freq) * 1000#
End Function

' FNV-1a 32 of b(0 .. n - 1) (Long multiplication wraps: overflow checks are off)
Private Function Fnv(b() As Byte, ByVal n As Long) As Long
    Dim h As Long, i As Long
    h = &H811C9DC5
    For i = 0 To n - 1
        h = (h Xor b(i)) * 16777619
    Next
    Fnv = h
End Function

Private Function BaseName(ByVal p As String) As String
    Dim k As Long
    k = InStrRev(p, "\")
    If InStrRev(p, "/") > k Then k = InStrRev(p, "/")
    BaseName = Mid$(p, k + 1)
End Function

Sub Main()
    Dim argv() As String, argc As Long, fmt As Long, outdir As String, listfile As String, nowrite As Boolean
    Dim lines() As String, li As Long, path As String, name As String, data() As Byte, n As Long
    Dim t As ktx2_transcoder, info As ktx2_image_level_info, levels As Long, level As Long, ok As Boolean
    Dim total As Long, cnt As Long, outb() As Byte, at As Long, t0 As Double, ms As Double, alltime As Double
    Dim res As Integer, f As Integer, txt As String, initms As Double, bpp As Long, outsize As Long

    argc = GetArgs(argv)
    If argc < 4 Then ExitProcess 2
    fmt = CLng(argv(1))
    outdir = argv(2)
    listfile = argv(3)
    nowrite = (argc > 4)
    If argc > 4 Then nowrite = (argv(4) = "nowrite")

    res = FreeFile
    Open outdir & "\results.txt" For Output As #res

    t0 = Now_()
    basisu_transcoder_init
    initms = Now_() - t0

    f = FreeFile
    Open listfile For Input As #f
    txt = Input$(LOF(f), #f)
    Close #f
    lines = Split(Replace(txt, vbCr, ""), vbLf)

    bpp = basis_get_bytes_per_block_or_pixel_(fmt)
    For li = 0 To UBound(lines)
        path = lines(li)
        If Len(path) = 0 Then GoTo next_file
        name = BaseName(path)
        n = ReadBytes(path, data)

        t0 = Now_()
        Set t = New ktx2_transcoder
        Set info = New ktx2_image_level_info
        ok = t.init(data, n)
        If ok Then ok = t.start_transcoding()
        levels = 0
        If ok Then levels = t.get_levels()
        ' the output's size: the levels' sizes added up first
        outsize = 0
        For level = 0 To levels - 1
            If Not t.get_image_level_info(info, level, 0, 0) Then ok = False: Exit For
            If basis_transcoder_format_is_uncompressed_(fmt) Then
                cnt = info.m_orig_width * info.m_orig_height
            Else
                cnt = info.m_total_blocks
            End If
            outsize = outsize + cnt * bpp
        Next
        If ok And outsize > 0 Then
            ReDim outb(0 To outsize - 1)
            at = 0
            For level = 0 To levels - 1
                t.get_image_level_info info, level, 0, 0
                If basis_transcoder_format_is_uncompressed_(fmt) Then
                    cnt = info.m_orig_width * info.m_orig_height
                Else
                    cnt = info.m_total_blocks
                End If
                If Not t.transcode_image_level(level, 0, 0, outb, at, cnt, fmt, 0) Then ok = False: Exit For
                at = at + cnt * bpp
            Next
        End If
        ms = Now_() - t0
        alltime = alltime + ms

        If Not ok Then
            Print #res, name & vbTab & "error" & vbTab & "transcode failed"
        Else
            Print #res, name & vbTab & t.get_width() & vbTab & t.get_height() & vbTab & levels & vbTab & outsize & vbTab & _
                LCase$(Right$("0000000" & Hex$(Fnv(outb, outsize)), 8)) & vbTab & Format$(ms, "0.00")
            If Not nowrite Then WriteBytes outdir & "\" & name & "." & fmt, outb, outsize
        End If
        Set t = Nothing
next_file:
    Next
    Print #res, "total" & vbTab & Format$(alltime, "0.00") & vbTab & Format$(initms, "0.00")
    Close #res
    ExitProcess 0
End Sub
