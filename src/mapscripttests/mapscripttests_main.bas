Attribute VB_Name = "mapscripttests_main"
' MapScriptTests.exe (VB-only, not part of the game): script-generated maps
' through WzMap.dll's Map_loadFromPath with a seed, dumped as
' tools/mapscriptref.cpp dumps them (the C++ library with QuickJS), for
' tools/map_script_test.py to compare.
'
'   MapScriptTests <command file> <output folder>
'
' Each line: "folder|maxplayers|seed" <TAB> output file (relative to the
' output folder); seed is the uint32_t's decimal value. A line "fpu" <TAB>
' output file writes the x87's control word as this VB process has it.
Option Explicit

Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, prgvt As Integer, prgpvarg As Long, pvargResult As Variant) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private outParts() As String         ' the output: its pieces, joined at the end
Private outCount As Long

Public Sub Put_(ByVal s As String)
    If outCount = 0 Then
        ReDim outParts(0 To 1023)
    ElseIf outCount > UBound(outParts) Then
        ReDim Preserve outParts(0 To 2 * outCount - 1)
    End If
    outParts(outCount) = s
    outCount = outCount + 1
End Sub

Public Function EndsWithLF() As Boolean
    If outCount > 0 Then EndsWithLF = (Right$(outParts(outCount - 1), 1) = vbLf)
End Function

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

Private Sub WriteText(ByVal path As String, ByVal s As String)
    Dim f As Integer, b() As Byte, n As Long
    n = utf8_encode(s, b)
    f = FreeFile
    If Len(Dir$(path)) > 0 Then Kill path
    Open path For Binary Access Write As #f
    If n > 0 Then Put #f, , b
    Close #f
End Sub

' msvcrt's _controlfp(0, 0): the x87's control word (a cdecl function: through DispCallFunc, CC_CDECL)
Private Function controlWord() As Long
    Dim a(0 To 1) As Variant, t(0 To 1) As Integer, p(0 To 1) As Long, r As Variant
    a(0) = 0&
    a(1) = 0&
    t(0) = vbLong
    t(1) = vbLong
    p(0) = VarPtr(a(0))
    p(1) = VarPtr(a(1))
    DispCallFunc 0, GetProcAddress(LoadLibraryA("msvcrt"), "_controlfp"), 1, vbLong, 2, t(0), p(0), r
    controlWord = r
End Function

' std::to_string(uint32_t)
Private Function UStr(ByVal v As Long) As String
    UStr = CStr(U32(v))
End Function

Private Sub dumpMap(ByVal arg As String)
    Dim parts() As String, logger As MapScriptLogger, m As Map, data As MapData, seed As Long, gotSeed As Long
    Dim i As Long, n As Long, v As ObjectVector, s As Structure, d As Droid, f As Feature, sep As String
    Dim crcTiles As Long, crcStructures As Long, crcDroids As Long, crcFeatures As Long
    parts = Split(arg, "|")
    Set logger = New MapScriptLogger
    seed = FromU32(CDbl(parts(2)))
    Set m = Map_loadFromPath(parts(0), MapType_SKIRMISH, CLng(parts(1)), seed, logger)
    If m Is Nothing Then
        Put_ "loadFromPath failed" & vbLf
        Exit Sub
    End If
    If Not m.scriptGeneratedMapSeed(gotSeed) Then gotSeed = 0
    Put_ "scriptgenerated " & IIf(m.wasScriptGenerated(), "1", "0") & " seed " & UStr(gotSeed) & vbLf
    Set data = m.mapData()
    If Not data Is Nothing Then
        n = data.mMapTiles_size()
        Put_ va("mapdata %u %u tiles %zu gateways %zu\n", data.width, data.height, n, data.mGateways_size())
        For i = 0 To n - 1
            If i Mod 16 = 15 Or i + 1 = n Then sep = vbLf Else sep = " "
            Put_ CStr(data.mMapTiles_height(i)) & ":" & CStr(data.mMapTiles_texture(i)) & sep
        Next
    Else
        Put_ "mapdata none" & vbLf
    End If
    Set v = m.mapStructures()
    If Not v Is Nothing Then
        Put_ va("structures %zu\n", v.size())
        For i = 0 To v.size() - 1
            Set s = v.at(i)
            Put_ "s " & s.name_ & " " & UStr(s.position_x) & " " & UStr(s.position_y) & " " & CStr(s.direction) & " " & _
                IIf(s.hasId, UStr(s.id), "-") & " " & CStr(s.player) & " " & CStr(s.modules) & vbLf
        Next
    Else
        Put_ "structures none" & vbLf
    End If
    Set v = m.mapDroids()
    If Not v Is Nothing Then
        Put_ va("droids %zu\n", v.size())
        For i = 0 To v.size() - 1
            Set d = v.at(i)
            Put_ "d " & d.name_ & " " & UStr(d.position_x) & " " & UStr(d.position_y) & " " & CStr(d.direction) & " " & _
                IIf(d.hasId, UStr(d.id), "-") & " " & CStr(d.player) & vbLf
        Next
    Else
        Put_ "droids none" & vbLf
    End If
    Set v = m.mapFeatures()
    If Not v Is Nothing Then
        Put_ va("features %zu\n", v.size())
        For i = 0 To v.size() - 1
            Set f = v.at(i)
            Put_ "f " & f.name_ & " " & UStr(f.position_x) & " " & UStr(f.position_y) & " " & CStr(f.direction) & " " & _
                IIf(f.hasId, UStr(f.id), "-") & " " & IIf(f.hasPlayer, CStr(f.player), "-") & vbLf
        Next
    Else
        Put_ "features none" & vbLf
    End If
    crcTiles = m.crcSumMapTiles(0)
    crcStructures = m.crcSumStructures(0)
    crcDroids = m.crcSumDroids(0)
    crcFeatures = m.crcSumFeatures(0)
    Put_ va("crc %08x %08x %08x %08x\n", crcTiles, crcStructures, crcDroids, crcFeatures)
End Sub

Sub Main()
    Dim argv() As String, argc As Long, lines() As String, i As Long, parts() As String
    Dim b() As Byte, n As Long, root As String, text As String
    argc = GetArgs(argv)
    If argc < 3 Then ExitProcess 2
    n = ReadBytes(argv(1), b)
    lines = Split(Replace(utf8_decode(b, 0, n), vbCr, ""), vbLf)
    root = argv(2)
    For i = 0 To UBound(lines)
        parts = Split(lines(i), vbTab)
        If UBound(parts) = 1 Then
            outCount = 0
            If parts(0) = "fpu" Then Put_ "controlfp " & Hex$(controlWord()) & vbLf Else dumpMap parts(0)
            text = ""
            If outCount > 0 Then
                ReDim Preserve outParts(0 To outCount - 1)
                text = Join(outParts, "")
            End If
            outCount = 0
            WriteText root & "\" & Replace(parts(1), "/", "\"), text
        End If
    Next
    ExitProcess 0
End Sub
