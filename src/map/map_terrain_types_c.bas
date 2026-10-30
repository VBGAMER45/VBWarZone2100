Attribute VB_Name = "map_terrain_types_c"
' lib/wzmaplib/src/map_terrain_types.cpp: ttypes.ttp, the terrain type of
' each tile texture. std::unique_ptr<TerrainTypeData> is the object or
' Nothing. A uint32_t is a Long with its bits (unsigned compares: ULT).
Option Explicit

' Arbitrary maximum number of terrain textures - used in look up table for terrain type
Private Const MAX_TILE_TEXTURES As Long = 255

Public Function loadTerrainTypes(ByVal filename As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As TerrainTypeData
    Dim path As String, pStream As BinaryIOStream, result As TerrainTypeData
    Dim aFileType(0 To 3) As Byte, n As Long, ok As Boolean
    Dim version As Long, quantity As Long, i As Long, pType As Long
    path = filename
    Set pStream = mapIO.openBinaryStream(filename, OpenMode_READ)
    If pStream Is Nothing Then
        ' Failed to open file
        Exit Function
    End If

    printLog_ pCustomLogger, 42, LogLevel_Info, "loadTerrainTypes", "Loading: %s", path

    Set result = New TerrainTypeData

    version = 0
    quantity = 0
    ok = pStream.readBytes(aFileType, 0, 4, n)
    If ok Then ok = (n = 4)
    If ok Then ok = (aFileType(0) = 116 And aFileType(1) = 116 And aFileType(2) = 121 And aFileType(3) = 112)  ' "ttyp"
    If ok Then ok = BinaryIOStream_readULE32(pStream, version)
    If ok Then ok = BinaryIOStream_readULE32(pStream, quantity)
    If Not ok Then
        printLog_ pCustomLogger, 56, LogLevel_Error, "loadTerrainTypes", "%s: Bad header", path
        Exit Function
    End If

    If ULT(version, 7) Or ULT(39, version) Then
        printLog_ pCustomLogger, 62, LogLevel_Error, "loadTerrainTypes", "%s: Unsupported binary ttp file version: %u", path, version
        Exit Function
    End If

    If Not ULT(quantity, MAX_TILE_TEXTURES) Then
        ' Workaround for fugly map editor bug, since we can't fix the map editor
        quantity = MAX_TILE_TEXTURES - 1
    End If

    For i = 0 To quantity - 1
        pType = 0
        If Not BinaryIOStream_readULE16(pStream, pType) Then
            ' Failed to read value
            printLog_ pCustomLogger, 78, LogLevel_Error, "loadTerrainTypes", "%s: Failed to read value # %u", path, i
            Exit Function
        End If
        If pType > TER_MAX Then
            printLog_ pCustomLogger, 84, LogLevel_Error, "loadTerrainTypes", "%s: Terrain type #%u (value: %u) out of range", path, i, pType
            Exit Function
        End If
        result.terrainTypes_push_back pType
    Next

    ' Check: extra bytes at end
    If Not pStream.endOfStream() Then
        printLog_ pCustomLogger, 94, LogLevel_Warning, "loadTerrainTypes", "%s: Unexpectedly did not reach end of stream - data may be corrupted", path
    End If

    Set loadTerrainTypes = result
End Function

Public Function writeTerrainTypes(ByVal ttypeData As TerrainTypeData, ByVal filename As String, ByVal mapIO As IOProvider, ByVal format As OutputFormat, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim path As String, pStream As BinaryIOStream, fileVersion As Long, numTtypes As Long
    Dim aFileType(0 To 3) As Byte, n As Long, ok As Boolean, i As Long, pType As Long
    path = filename
    Set pStream = mapIO.openBinaryStream(filename, OpenMode_WRITE)
    If pStream Is Nothing Then
        printLog_ pCustomLogger, 107, LogLevel_Error, "writeTerrainTypes", "Failed to open file for writing: %s", path
        Exit Function
    End If

    printLog_ pCustomLogger, 111, LogLevel_Info, "writeTerrainTypes", "Writing: %s", path

    fileVersion = 8 ' the format hasn't really changed
    Select Case format
    Case OutputFormat_VER1_BINARY_OLD
        fileVersion = 8     ' flaME expects version 8
    Case OutputFormat_VER2
        fileVersion = 8     ' stick with version 8 for now
    Case OutputFormat_VER3
        fileVersion = 39    ' use version 39 (the last version before maplib refactor) for now
    End Select

    numTtypes = ttypeData.terrainTypes_size()

    ' write header
    aFileType(0) = 116  ' t
    aFileType(1) = 116  ' t
    aFileType(2) = 121  ' y
    aFileType(3) = 112  ' p
    ok = pStream.writeBytes(aFileType, 0, 4, n)
    If ok Then ok = (n = 4)
    If ok Then ok = BinaryIOStream_writeULE32(pStream, fileVersion)
    If ok Then ok = BinaryIOStream_writeULE32(pStream, numTtypes)
    If Not ok Then
        printLog_ pCustomLogger, 138, LogLevel_Error, "writeTerrainTypes", "Failed writing header to: %s", path
        Exit Function
    End If

    For i = 0 To numTtypes - 1
        pType = ttypeData.terrainTypes(i) And &HFFFF&
        If Not BinaryIOStream_writeULE16(pStream, pType) Then
            printLog_ pCustomLogger, 148, LogLevel_Error, "writeTerrainTypes", "Error writing terrain type data to: %s", path
            Exit Function
        End If
    Next

    writeTerrainTypes = True
End Function
