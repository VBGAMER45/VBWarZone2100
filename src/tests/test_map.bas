Attribute VB_Name = "test_map"
' WzTests: wzmaplib (WzMap.dll) against tools/mapref.cpp, the same commands
' through the real library.
'   map     "folder|maptype|maxplayers": Map_loadFromPath, then a dump of
'           the loaded format, the map data, the objects, the terrain
'           types, the CRCs, and every log line (CaptureLogger)
'   export  "folder|maptype|maxplayers|format": the map loaded and written
'           by Map_exportMapToPath into <root>/<output> (a folder that
'           tools/run_tests.py makes); the output is the log
Option Explicit

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

Private Function mapTypeOf(ByVal s As String) As MapType
    If s = "campaign" Then
        mapTypeOf = MapType_CAMPAIGN
    ElseIf s = "savegame" Then
        mapTypeOf = MapType_SAVEGAME
    Else
        mapTypeOf = MapType_SKIRMISH
    End If
End Function

' std::to_string(uint32_t)
Private Function UStr(ByVal v As Long) As String
    UStr = CStr(U32(v))
End Function

Private Sub dumpMap(ByVal arg As String)
    Dim parts() As String, logger As CaptureLogger, m As Map, format As LoadedFormat, data As MapData
    Dim i As Long, n As Long, v As ObjectVector, s As Structure, d As Droid, f As Feature, ttypes As TerrainTypeData
    Dim crcTiles As Long, crcStructures As Long, crcDroids As Long, crcFeatures As Long, sep As String
    parts = Split(arg, "|")
    Set logger = New CaptureLogger
    Set m = Map_loadFromPath(parts(0), mapTypeOf(parts(1)), CLng(parts(2)), 0, logger)
    If m Is Nothing Then
        Put_ "loadFromPath failed" & vbLf
        Exit Sub
    End If
    If m.loadedMapFormat(format) Then Put_ va("format %d\n", format) Else Put_ "format -1" & vbLf
    Set data = m.mapData()
    If Not data Is Nothing Then
        n = data.mMapTiles_size()
        Put_ va("mapdata %u %u tiles %zu gateways %zu\n", data.width, data.height, n, data.mGateways_size())
        For i = 0 To n - 1
            If i Mod 16 = 15 Or i + 1 = n Then sep = vbLf Else sep = " "
            Put_ CStr(data.mMapTiles_height(i)) & ":" & CStr(data.mMapTiles_texture(i)) & sep
        Next
        For i = 0 To data.mGateways_size() - 1
            Put_ va("gateway %u %u %u %u\n", data.mGateways_x1(i), data.mGateways_y1(i), data.mGateways_x2(i), data.mGateways_y2(i))
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
    Set ttypes = m.mapTerrainTypes()
    If Not ttypes Is Nothing Then
        n = ttypes.terrainTypes_size()
        Put_ va("terraintypes %zu\n", n)
        For i = 0 To n - 1
            If i Mod 32 = 31 Or i + 1 = n Then sep = vbLf Else sep = " "
            Put_ CStr(ttypes.terrainTypes(i)) & sep
        Next
    Else
        Put_ "terraintypes none" & vbLf
    End If
    ' one at a time: a part that failed to load is loaded again (and logs again), in this order
    crcTiles = m.crcSumMapTiles(0)
    crcStructures = m.crcSumStructures(0)
    crcDroids = m.crcSumDroids(0)
    crcFeatures = m.crcSumFeatures(0)
    Put_ va("crc %08x %08x %08x %08x\n", crcTiles, crcStructures, crcDroids, crcFeatures)
End Sub

Private Sub exportMap(ByVal arg As String, ByVal outFolder As String)
    Dim parts() As String, logger As CaptureLogger, m As Map, mapType_ As MapType, maxPlayers As Long, ok As Boolean
    parts = Split(arg, "|")
    Set logger = New CaptureLogger
    mapType_ = mapTypeOf(parts(1))
    maxPlayers = CLng(parts(2))
    Set m = Map_loadFromPath(parts(0), mapType_, maxPlayers, 0, logger)
    If m Is Nothing Then
        Put_ "loadFromPath failed" & vbLf
        Exit Sub
    End If
    ok = Map_exportMapToPath(m, outFolder, mapType_, maxPlayers, CLng(parts(3)), logger)
    If ok Then Put_ "export ok" & vbLf Else Put_ "export failed" & vbLf
End Sub

' The output of one command (mapref's out)
Public Function RunMap(ByVal mode As String, ByVal arg As String, ByVal root As String, ByVal outName As String) As String
    outCount = 0
    If mode = "map" Then
        dumpMap arg
    ElseIf mode = "export" Then
        exportMap arg, root & "/" & outName
    End If
    If outCount > 0 Then
        ReDim Preserve outParts(0 To outCount - 1)
        RunMap = Join(outParts, "")
    End If
    outCount = 0
End Function
