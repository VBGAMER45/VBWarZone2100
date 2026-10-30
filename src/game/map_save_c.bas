Attribute VB_Name = "map_save_c"
' src/map.cpp: its save game parts (map_c has the rest of map.cpp): mapSaveToWzMapData,
' writeVisibilityData and readVisibilityData. A visibility plane is written and read as one block of
' bytes (the C++ writes and reads them one PHYSFS_writeUBE8 / PHYSFS_readUBE8 at a time through
' WZ_PHYSFS_SETBUFFER's 4096-byte buffer: the same file).
Option Explicit

Private Const CURRENT_VERSION_NUM As Long = 39  ' gamedef.h: VERSION_39

Public Function mapSaveToWzMapData(ByVal output As MapData) As Boolean
    Dim numMapTiles As Long, i As Long, height_ As Long, gwCount As Long, gw As GATEWAY
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long

    output.width = mapWidth
    output.height = mapHeight

    ' Write out the map tile data
    numMapTiles = output.width * output.height
    output.mMapTiles_clear
    output.mMapTiles_reserve numMapTiles
    For i = 0 To numMapTiles - 1
        ' WzMap::MapData::MapTile mapDataTile = {};
        If terrainType(i) = TER_WATER Then
            height_ = (psMapTiles(i).waterLevel + world_coord(1) \ 3) ' this magic number stuff should match afterMapLoad()'s handling of water tiles (??)
        Else
            height_ = psMapTiles(i).height
        End If
        output.mMapTiles_push_back height_ And &HFFFF&, psMapTiles(i).texture
    Next

    ' Write out the gateway data
    output.mGateways_clear
    gwCount = gwNumGateways()
    output.mGateways_reserve gwCount
    For i = 0 To gwCount - 1
        gwGetGateway i, gw
        ' WzMap::MapData::Gateway gw = {};
        x1 = gw.x1
        y1 = gw.y1
        x2 = gw.x2
        y2 = gw.y2
        If Not (x1 = x2 Or y1 = y2) Then ASSERT_FAILURE "gw.x1 == gw.x2 || gw.y1 == gw.y2", "src/map.cpp:1243", "mapSaveToWzMapData", "Invalid gateway coordinates (%d, %d, %d, %d)", _
               x1, y1, x2, y2
        If Not (x1 < mapWidth And y1 < mapHeight And x2 < mapWidth And y2 < mapHeight) Then ASSERT_FAILURE "gw.x1 < mapWidth && gw.y1 < mapHeight && gw.x2 < mapWidth && gw.y2 < mapHeight", "src/map.cpp:1245", "mapSaveToWzMapData", _
               "Bad gateway dimensions for savegame"
        output.mGateways_push_back x1, y1, x2, y2
    Next

    mapSaveToWzMapData = True
End Function

' -----------------------------------------------------------------------------------
' This will save out the visibility data
Public Function writeVisibilityData(ByVal fileName As String) As Boolean
    Dim fileHandle As Long, b() As Byte, planes As Long, plane As Long, i As Long, n As Long, shift As Long

    fileHandle = openSaveFile(fileName)
    If fileHandle = 0 Then
        writeVisibilityData = False
        Exit Function
    End If

    ' (WZ_PHYSFS_SETBUFFER(fileHandle, 4096): the planes are written whole)

    ' VIS_SAVEHEADER fileHeader: 'v' 'i' 's' 'd', CURRENT_VERSION_NUM
    ReDim b(0 To 3)
    b(0) = Asc("v")
    b(1) = Asc("i")
    b(2) = Asc("s")
    b(3) = Asc("d")

    ' Write out the current file header
    If WZ_PHYSFS_writeBytes(fileHandle, b, 0, 4) <> 4 Then GoTo headerFailed
    If PHYSFS_writeUBE32(fileHandle, CURRENT_VERSION_NUM) = 0 Then GoTo headerFailed

    planes = (game.maxPlayers + 7) \ 8

    n = mapWidth * mapHeight
    If n > 0 Then ReDim b(0 To n - 1)
    For plane = 0 To planes - 1
        shift = plane * 8
        For i = 0 To n - 1
            ' psMapTiles[i].tileExploredBits >> (plane * 8), as a uint8_t
            b(i) = (psMapTiles(i).tileExploredBits \ (2 ^ shift)) And &HFF&
        Next
        If n > 0 Then
            If WZ_PHYSFS_writeBytes(fileHandle, b, 0, n) <> n Then
                debug_ 1796, LOG_ERROR, "writeVisibilityData", "writeVisibilityData: could not write to %s; PHYSFS error: %s", fileName, WZ_PHYSFS_getLastError()
                PHYSFS_close fileHandle
                writeVisibilityData = False
                Exit Function
            End If
        End If
    Next

    ' Everything is just fine!
    PHYSFS_close fileHandle
    writeVisibilityData = True
    Exit Function

headerFailed:
    debug_ 1783, LOG_ERROR, "writeVisibilityData", "writeVisibilityData: could not write header to %s; PHYSFS error: %s", fileName, WZ_PHYSFS_getLastError()
    PHYSFS_close fileHandle
    writeVisibilityData = False
End Function

' -----------------------------------------------------------------------------------
' This will read in the visibility data
Public Function readVisibilityData(ByVal fileName As String) As Boolean
    Dim fileHandle As Long, aFileType(0 To 3) As Byte, b() As Byte, version As Long
    Dim expectedFileSize As Double, fileSize As Double, i As Long, planes As Long, plane As Long, n As Long

    fileHandle = openLoadFile(fileName, False)
    If fileHandle = 0 Then
        ' Failure to open the file is no failure to read it
        readVisibilityData = True
        Exit Function
    End If

    ' Read the header from the file
    ReDim b(0 To 3)
    If WZ_PHYSFS_readBytes(fileHandle, b, 0, 4) <> 4 Then GoTo headerFailed
    For i = 0 To 3
        aFileType(i) = b(i)
    Next
    If PHYSFS_readUBE32(fileHandle, version) = 0 Then GoTo headerFailed

    ' Check the header to see if we've been given a file of the right type
    If aFileType(0) <> Asc("v") Or aFileType(1) <> Asc("i") Or aFileType(2) <> Asc("s") Or aFileType(3) <> Asc("d") Then
        debug_ 1838, LOG_ERROR, "readVisibilityData", "readVisibilityData: Weird file type found? Has header letters - '%c' '%c' '%c' '%c' (should be 'v' 'i' 's' 'd')", _
              CLng(aFileType(0)), _
              CLng(aFileType(1)), _
              CLng(aFileType(2)), _
              CLng(aFileType(3))

        PHYSFS_close fileHandle
        readVisibilityData = False
        Exit Function
    End If

    planes = (game.maxPlayers + 7) \ 8

    ' Validate the filesize
    n = mapWidth * mapHeight
    expectedFileSize = 4 + 4 + CDbl(n) * planes    ' sizeof(fileHeader.aFileType) + sizeof(fileHeader.version) + mapWidth * mapHeight * planes
    fileSize = PHYSFS_fileLength(fileHandle)
    If fileSize <> expectedFileSize Then
        PHYSFS_close fileHandle
        ASSERT_FAILURE "!""readVisibilityData: unexpected filesize""", "src/map.cpp:1856", "readVisibilityData", "readVisibilityData: unexpected filesize; should be %u, but is %u", expectedFileSize, fileSize

        readVisibilityData = False
        Exit Function
    End If

    ' For every tile...
    For i = 0 To n - 1
        psMapTiles(i).tileExploredBits = 0
    Next
    If n > 0 Then ReDim b(0 To n - 1)
    For plane = 0 To planes - 1
        ' Get the visibility data
        If n > 0 Then
            If WZ_PHYSFS_readBytes(fileHandle, b, 0, n) <> n Then
                debug_ 1874, LOG_ERROR, "readVisibilityData", "readVisibilityData: could not read from %s; PHYSFS error: %s", fileName, WZ_PHYSFS_getLastError()
                PHYSFS_close fileHandle
                readVisibilityData = False
                Exit Function
            End If
        End If
        For i = 0 To n - 1
            psMapTiles(i).tileExploredBits = psMapTiles(i).tileExploredBits Or (CLng(b(i)) * (2 ^ (plane * 8)))
        Next
    Next

    ' Close the file
    PHYSFS_close fileHandle

    ' Hopefully everything's just fine by now
    readVisibilityData = True
    Exit Function

headerFailed:
    debug_ 1827, LOG_ERROR, "readVisibilityData", "readVisibilityData: error while reading header from file: %s", WZ_PHYSFS_getLastError()
    PHYSFS_close fileHandle
    readVisibilityData = False
End Function
