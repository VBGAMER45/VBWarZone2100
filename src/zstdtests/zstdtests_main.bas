Attribute VB_Name = "zstdtests_main"
' ZstdTests.exe (VB-only, not part of the game; cluster T, fork T1): decompresses zstd frames with the port of
' zstd 1.4.9's decoder (src/basis/zstd_*.bas, compiled in) and with the same decoder in C as the 32-bit build has
' it (z32.dll beside the exe: ref/basis_universal/zstd/zstddeclib.c compiled by gcc -m32, tools/zstdref), and
' compares the two: the result (a size or an error code) and the bytes regenerated. tools/zstd_test.py makes the
' frames and checks the outputs' hashes against the 64-bit C library. Built twice: ZstdTests (array bounds
' checked) and ZstdTestsFast (optimised as WzBasis: no bounds checks).
'
'   ZstdTests <list file> <output folder>
'
' The list: one line per frame, "name TAB file TAB offset TAB length TAB dstCapacity". The destination is a
' buffer with 64 guard bytes before and after the capacity (a write outside it is reported).
' Output: <output folder>\results.txt, a line per frame:
'   name TAB VB result TAB C result TAB same (1/0) TAB guards intact (1/0) TAB zlib's CRC-32 of the VB output TAB
'   VB microseconds TAB C microseconds
' and a last line "total TAB frames TAB differences TAB VB bytes out TAB VB ms TAB C ms".
Option Explicit

Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function z32_decompress Lib "z32.dll" (dst As Byte, ByVal dstCapacity As Long, src As Byte, ByVal srcSize As Long) As Long

Private Const GUARD As Long = 64
Private Const GUARD_BYTE As Byte = &HCD

Private crcTable(0 To 255) As Long

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

Private Sub InitCrc()
    Dim n As Long, c As Long, k As Long
    For n = 0 To 255
        c = n
        For k = 1 To 8
            If (c And 1) <> 0 Then
                c = &HEDB88320 Xor (((c And &HFFFFFFFE) \ 2) And &H7FFFFFFF)
            Else
                c = ((c And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next k
        crcTable(n) = c
    Next n
End Sub

' zlib's crc32 of b(first .. first + n - 1)
Private Function Crc32(b() As Byte, ByVal first As Long, ByVal n As Long) As Long
    Dim c As Long, i As Long
    c = -1
    For i = first To first + n - 1
        c = crcTable((c Xor b(i)) And &HFF&) Xor (((c And &HFFFFFF00) \ &H100&) And &HFFFFFF)
    Next i
    Crc32 = Not c
End Function

Private Function GuardsIntact(b() As Byte, ByVal capacity As Long) As Boolean
    Dim i As Long
    For i = 0 To GUARD - 1
        If b(i) <> GUARD_BYTE Then Exit Function
        If b(GUARD + capacity + i) <> GUARD_BYTE Then Exit Function
    Next i
    GuardsIntact = True
End Function

Private Function Hex8(ByVal v As Long) As String
    Hex8 = Right$("0000000" & Hex$(v), 8)
End Function

Public Sub Main()
    Dim argv() As String, argc As Long, lines() As String, listText As String, f As Integer, o As Integer
    Dim i As Long, parts() As String, name As String, path As String, offset As Long, length As Long, capacity As Long
    Dim curPath As String, data() As Byte, dataLen As Long, src() As Byte
    Dim dstVB() As Byte, dstC() As Byte, rVB As Long, rC As Long, same As Boolean, guards As Boolean, k As Long
    Dim freq As Currency, t0 As Currency, t1 As Currency, t2 As Currency, usVB As Double, usC As Double
    Dim frames As Long, diffs As Long, bytesOut As Double, msVB As Double, msC As Double

    argc = GetArgs(argv)
    If argc < 3 Then ExitProcess 2
    QueryPerformanceFrequency freq
    InitCrc

    f = FreeFile
    Open argv(1) For Binary Access Read As #f
    listText = Space$(LOF(f))
    Get #f, , listText
    Close #f
    lines = Split(Replace(listText, vbCr, ""), vbLf)

    o = FreeFile
    Open argv(2) & "\results.txt" For Output As #o
    For i = 0 To UBound(lines)
        If Len(lines(i)) = 0 Then GoTo nextLine
        parts = Split(lines(i), vbTab)
        name = parts(0)
        path = parts(1)
        offset = CLng(parts(2))
        length = CLng(parts(3))
        capacity = CLng(parts(4))
        If path <> curPath Then
            dataLen = ReadBytes(path, data)
            curPath = path
        End If

        ' the source: the frame where it is in its file (VB), and a copy of it for the C (a pointer to its bytes)
        ReDim src(0 To length + 16)
        If length > 0 Then CopyMemory src(0), data(offset), length

        ReDim dstVB(0 To capacity + 2 * GUARD - 1)
        ReDim dstC(0 To capacity + 2 * GUARD - 1)
        For k = 0 To capacity + 2 * GUARD - 1
            dstVB(k) = GUARD_BYTE
            dstC(k) = GUARD_BYTE
        Next k

        QueryPerformanceCounter t0
        rVB = ZSTD_decompress(dstVB, GUARD, capacity, data, offset, length)
        QueryPerformanceCounter t1
        rC = z32_decompress(dstC(GUARD), capacity, src(0), length)
        QueryPerformanceCounter t2
        usVB = (t1 - t0) / freq * 1000000#
        usC = (t2 - t1) / freq * 1000000#
        msVB = msVB + usVB / 1000#
        msC = msC + usC / 1000#

        same = (rVB = rC)
        If same And Not ZSTD_isError(rVB) Then
            If same And rVB > 0 Then
                For k = 0 To rVB - 1
                    If dstVB(GUARD + k) <> dstC(GUARD + k) Then same = False: Exit For
                Next k
            End If
            bytesOut = bytesOut + rVB
        End If
        guards = GuardsIntact(dstVB, capacity)
        frames = frames + 1
        If Not same Or Not guards Then diffs = diffs + 1

        If ZSTD_isError(rVB) Then
            Print #o, name & vbTab & rVB & vbTab & rC & vbTab & IIf(same, 1, 0) & vbTab & IIf(guards, 1, 0) & vbTab & "-" & vbTab & Format$(usVB, "0") & vbTab & Format$(usC, "0")
        Else
            Print #o, name & vbTab & rVB & vbTab & rC & vbTab & IIf(same, 1, 0) & vbTab & IIf(guards, 1, 0) & vbTab & Hex8(Crc32(dstVB, GUARD, rVB)) & vbTab & Format$(usVB, "0") & vbTab & Format$(usC, "0")
        End If
nextLine:
    Next i
    Print #o, "total" & vbTab & frames & vbTab & diffs & vbTab & Format$(bytesOut, "0") & vbTab & Format$(msVB, "0.0") & vbTab & Format$(msC, "0.0")
    Close #o
    ExitProcess 0
End Sub
