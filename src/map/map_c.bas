Attribute VB_Name = "map_c"
' lib/wzmaplib/src/map.cpp: loading and writing a map's files (game.map, and
' the structures, droids and features in the old binary .bjo files or in
' JSON), and the static members of class Map (Map_loadFromPath,
' Map_exportMapToPath, Map_expectedFileNames). The rest of class Map is
' Map.cls, which calls the loaders here (static in map.cpp, Public here).
' * A uint32_t is a Long with its 32 bits: unsigned compares through ULT,
'   %u prints it unsigned; uint8_t, uint16_t and int8_t values are Longs.
' * optional<FileLoadResult<T>> is the Boolean result, with the objects (an
'   ObjectVector) and the file's format version in ByRef parameters;
'   optional<MapDataLoadResult> likewise; struct JsonParsingContext
'   {filename, jsonPath} is its two strings.
' * The templates over Structure, Droid and Feature are one late-bound copy
'   (mapObj As Object): the three classes share the members they use.
'   jsonGetListOfType<T> takes T as its bit count (32: uint32_t, 16:
'   uint16_t).
' * try / catch (const std::exception &e) around JSON code: the Try*
'   helpers, which return false with e.what() (WzFramework's json errors).
' * JSON written is ordered_json (insertion order); JSON read is json.
' * The C++ bugs are kept: leftPadStrToMinimumLength pads on the right (so
'   v1 JSON keys can repeat, and the later object replaces the earlier);
'   jsonGetRootMapObjectsContainer assigns the container into the root
'   object, and is not given the logger.
' Missing: Map::calculateMapStats (map_stats.cpp). Script-generated maps:
' runMapScript (map_script_c).
Option Explicit

Private Const VERSION_39 As Long = 39            ' last game.map version in the old format (with bytes for tile height, must be multiplied by ELEVATION_SCALE)
Private Const VERSION_40 As Long = 40            ' game.map version with full-range tile height
Private Const CURRENT_VERSION_NUM As Long = VERSION_40

Private Const MAX_PLAYERS As Long = 11           ' Maximum number of players in the game.

Private Const SS_BUILT As Long = 1

Private Const MaxGameJSFileSize As Long = 1048576    ' 1 * 1024 * 1024

' 65536 / 360 = 8192 / 45, with a bit less overflow risk.
' DEG(degrees) ((degrees) * 8192 / 45), on a uint32_t (the multiplication
' wraps), stored into a uint16_t (shared fixedpoint_h's DEG is the plain int macro)
Private Function DEG_u16(ByVal degrees As Long) As Long
    DEG_u16 = Int(U32(FromU32(U32(degrees) * 8192#)) / 45#) And &HFFFF&
End Function


' x + 1 for a uint32_t
Private Function U32Inc(ByVal x As Long) As Long
    U32Inc = FromU32(U32(x) + 1#)
End Function

' The C string in buf: the bytes up to the first NUL, as UTF-8
Private Function CStringOf(buf() As Byte, ByVal n As Long) As String
    Dim i As Long
    For i = 0 To n - 1
        If buf(i) = 0 Then Exit For
    Next
    CStringOf = utf8_decode(buf, 0, i)
End Function

'--------------------------------------------------------------------------
' MARK: - Handling game.map / MapData files

Public Function loadMapData_Internal(ByVal filename As String, ByVal mapIO As IOProvider, ByVal pCustomLogger As LoggingProtocol, resultMapData As MapData, resultFileFormatVersion As Long) As Boolean
    Dim path As String, pStream As BinaryIOStream, map_ As MapData, mapVersion As Long
    Dim aFileType(0 To 3) As Byte, n As Long, ok As Boolean, w As Long, h As Long
    Dim numMapTiles As Long, i As Long, texture As Long, height_ As Long
    Dim gwVersion As Long, numGateways As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    path = filename
    Set pStream = mapIO.openBinaryStream(filename, OpenMode_READ)

    If pStream Is Nothing Then
        printLog_ pCustomLogger, 64, LogLevel_Error, "loadMapData_Internal", "%s not found", path
        Exit Function
    End If

    printLog_ pCustomLogger, 68, LogLevel_Info, "loadMapData_Internal", "Loading: %s", path

    Set map_ = New MapData
    mapVersion = 0
    ok = pStream.readBytes(aFileType, 0, 4, n)
    If ok Then ok = (n = 4)
    If ok Then ok = BinaryIOStream_readULE32(pStream, mapVersion)
    If ok Then ok = BinaryIOStream_readULE32(pStream, w)
    If ok Then ok = BinaryIOStream_readULE32(pStream, h)
    If ok Then ok = (aFileType(0) = 109 And aFileType(1) = 97 And aFileType(2) = 112)    ' "map"
    If Not ok Then
        printLog_ pCustomLogger, 81, LogLevel_Error, "loadMapData_Internal", "Bad header in %s", path
        Exit Function
    End If
    map_.width = w
    map_.height = h

    If Not ULT(9, mapVersion) Then
        printLog_ pCustomLogger, 87, LogLevel_Error, "loadMapData_Internal", "%s: Unsupported save format version %u", path, mapVersion
        Exit Function
    End If
    If ULT(CURRENT_VERSION_NUM, mapVersion) Then
        printLog_ pCustomLogger, 92, LogLevel_Error, "loadMapData_Internal", "%s: Undefined save format version %u", path, mapVersion
        Exit Function
    End If
    If U32(w) * U32(h) > MAP_MAXAREA Then
        printLog_ pCustomLogger, 97, LogLevel_Error, "loadMapData_Internal", "Map %s too large : %d %d", path, w, h
        Exit Function
    End If

    If Not ULT(1, w) Or Not ULT(1, h) Then
        printLog_ pCustomLogger, 103, LogLevel_Error, "loadMapData_Internal", "Map %s is too small : %u, %u", path, w, h
        Exit Function
    End If

    ' Load in the map data
    numMapTiles = w * h
    map_.mMapTiles_reserve numMapTiles
    If Not ULT(mapVersion, VERSION_40) Then
        ' load full-range map tile heights
        For i = 0 To numMapTiles - 1
            texture = 0
            height_ = 0
            ok = BinaryIOStream_readULE16(pStream, texture)
            If ok Then ok = BinaryIOStream_readULE16(pStream, height_)
            If Not ok Then
                printLog_ pCustomLogger, 118, LogLevel_Error, "loadMapData_Internal", "%s: Error during savegame load", path
                Exit Function
            End If
            If height_ > TILE_MAX_HEIGHT Then
                printLog_ pCustomLogger, 123, LogLevel_Error, "loadMapData_Internal", "%s: Tile height (%u) exceeds TILE_MAX_HEIGHT (%zu)", path, height_, TILE_MAX_HEIGHT
                Exit Function
            End If
            map_.mMapTiles_push_back height_, texture
        Next
    Else
        ' load old map tile heights where tile-heights fit into a byte (and raw value is divided by ELEVATION_SCALE)
        For i = 0 To numMapTiles - 1
            ok = BinaryIOStream_readULE16(pStream, texture)
            If ok Then ok = BinaryIOStream_readULE8(pStream, height_)
            If Not ok Then
                printLog_ pCustomLogger, 139, LogLevel_Error, "loadMapData_Internal", "%s: Error during savegame load", path
                Exit Function
            End If
            map_.mMapTiles_push_back height_ * ELEVATION_SCALE, texture
        Next
    End If

    ok = BinaryIOStream_readULE32(pStream, gwVersion)
    If ok Then ok = BinaryIOStream_readULE32(pStream, numGateways)
    If ok Then ok = (gwVersion = 1)
    If Not ok Then
        printLog_ pCustomLogger, 153, LogLevel_Error, "loadMapData_Internal", "Bad gateway in %s", path
        Exit Function
    End If

    If numGateways >= 0 And numGateways <= MAP_MAXAREA Then map_.mGateways_reserve numGateways   ' (a reservation only)
    i = 0
    Do While ULT(i, numGateways)
        x1 = 0: y1 = 0: x2 = 0: y2 = 0
        ok = BinaryIOStream_readULE8(pStream, x1)
        If ok Then ok = BinaryIOStream_readULE8(pStream, y1)
        If ok Then ok = BinaryIOStream_readULE8(pStream, x2)
        If ok Then ok = BinaryIOStream_readULE8(pStream, y2)
        If Not ok Then
            printLog_ pCustomLogger, 164, LogLevel_Error, "loadMapData_Internal", "%s: Failed to read gateway info", path
            Exit Function
        End If
        map_.mGateways_push_back x1, y1, x2, y2
        i = i + 1
    Loop

    Set resultMapData = map_
    resultFileFormatVersion = mapVersion
    loadMapData_Internal = True
End Function

' Load the map data
Public Function loadMapData(ByVal filename As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As MapData
    Dim result As MapData, version As Long
    If Not loadMapData_Internal(filename, mapIO, pCustomLogger, result, version) Then
        Exit Function
    End If
    Set loadMapData = result
End Function

' Write the map data
Public Function writeMapData(ByVal map_ As MapData, ByVal filename As String, ByVal mapIO As IOProvider, ByVal format As OutputFormat, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim path As String, pStream As BinaryIOStream, mapVersion As Long, aFileType(0 To 3) As Byte, n As Long, ok As Boolean
    Dim numMapTiles As Long, i As Long, tileHeight As Long, gwVersion As Long, numGateways As Long
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    path = filename
    Set pStream = mapIO.openBinaryStream(filename, OpenMode_WRITE)

    If pStream Is Nothing Then
        printLog_ pCustomLogger, 193, LogLevel_Error, "writeMapData", "Failed to open file for writing: %s", path
        Exit Function
    End If

    printLog_ pCustomLogger, 197, LogLevel_Info, "writeMapData", "Writing: %s", path

    mapVersion = CURRENT_VERSION_NUM
    Select Case format
    Case OutputFormat_VER1_BINARY_OLD
        mapVersion = 10
    Case OutputFormat_VER2
        mapVersion = 39
    Case OutputFormat_VER3
        mapVersion = 40
    End Select

    ' write header
    aFileType(0) = 109  ' m
    aFileType(1) = 97   ' a
    aFileType(2) = 112  ' p
    aFileType(3) = 32   ' (space)
    ok = pStream.writeBytes(aFileType, 0, 4, n)
    If ok Then ok = (n = 4)
    If ok Then ok = BinaryIOStream_writeULE32(pStream, mapVersion)
    If ok Then ok = BinaryIOStream_writeULE32(pStream, map_.width)
    If ok Then ok = BinaryIOStream_writeULE32(pStream, map_.height)
    If Not ok Then
        printLog_ pCustomLogger, 224, LogLevel_Error, "writeMapData", "Failed writing header to: %s", path
        Exit Function
    End If

    If U32(map_.width) * U32(map_.height) > MAP_MAXAREA Then
        printLog_ pCustomLogger, 230, LogLevel_Error, "writeMapData", "Map too large : %u x %u", map_.width, map_.height
        Exit Function
    End If
    If Not ULT(9, mapVersion) Then
        printLog_ pCustomLogger, 235, LogLevel_Error, "writeMapData", "Unsupported save format version %u", mapVersion
        Exit Function
    End If
    If Not ULT(1, map_.width) Or Not ULT(1, map_.height) Then
        printLog_ pCustomLogger, 240, LogLevel_Error, "writeMapData", "Map is too small : %u x %u", map_.width, map_.height
        Exit Function
    End If

    ' Write out the map tiles
    numMapTiles = map_.width * map_.height
    If numMapTiles <> map_.mMapTiles_size() Then
        printLog_ pCustomLogger, 248, LogLevel_Error, "writeMapData", "Map width x height (%u x %u) != number of map tiles (%zu)", map_.width, map_.height, map_.mMapTiles_size()
        Exit Function
    End If

    If Not ULT(mapVersion, VERSION_40) Then
        ' write new tile info, with full-range tile heights
        For i = 0 To numMapTiles - 1
            If map_.mMapTiles_height(i) > TILE_MAX_HEIGHT Then
                ' tile height exceeds maximum
                printLog_ pCustomLogger, 261, LogLevel_Error, "writeMapData", "Error writing maptile; tile height (%u) exceeds maximum tile height (%zu)", map_.mMapTiles_height(i), TILE_MAX_HEIGHT
                Exit Function
            End If
            ok = BinaryIOStream_writeULE16(pStream, map_.mMapTiles_texture(i))
            If ok Then ok = BinaryIOStream_writeULE16(pStream, map_.mMapTiles_height(i))
            If Not ok Then
                printLog_ pCustomLogger, 266, LogLevel_Error, "writeMapData", "Error writing maptile to: %s", path
                Exit Function
            End If
        Next
    Else
        ' write old tile info, where tile-heights fit into a byte (and raw value is divided by ELEVATION_SCALE)
        For i = 0 To numMapTiles - 1
            If map_.mMapTiles_height(i) > 255 * ELEVATION_SCALE Then
                ' tile height exceeds maximum supported by this old format
                printLog_ pCustomLogger, 280, LogLevel_Error, "writeMapData", "Error writing maptile; tile height (%u) exceeds maximum supported by output format (%zu)", map_.mMapTiles_height(i), 255
                Exit Function
            End If
            tileHeight = map_.mMapTiles_height(i) \ ELEVATION_SCALE
            ok = BinaryIOStream_writeULE16(pStream, map_.mMapTiles_texture(i))
            If ok Then ok = BinaryIOStream_writeULE8(pStream, tileHeight)
            If Not ok Then
                printLog_ pCustomLogger, 286, LogLevel_Error, "writeMapData", "Error writing maptile to: %s", path
                Exit Function
            End If
        Next
    End If

    ' Write the gateway header
    gwVersion = 1
    numGateways = map_.mGateways_size()
    ok = BinaryIOStream_writeULE32(pStream, gwVersion)
    If ok Then ok = BinaryIOStream_writeULE32(pStream, numGateways)
    If Not ok Then
        printLog_ pCustomLogger, 297, LogLevel_Error, "writeMapData", "Error writing gateway header to: %s", path
        Exit Function
    End If

    ' Write out the gateways
    For i = 0 To numGateways - 1
        x1 = map_.mGateways_x1(i)
        y1 = map_.mGateways_y1(i)
        x2 = map_.mGateways_x2(i)
        y2 = map_.mGateways_y2(i)
        If x1 <> x2 And y1 <> y2 Then
            printLog_ pCustomLogger, 307, LogLevel_Warning, "writeMapData", "Invalid gateway coordinates (%u, %u, %u, %u)", x1, y1, x2, y2
            ' for now, write anyway
        End If
        If ULT(map_.width, x1) Or ULT(map_.width, x2) Or ULT(map_.height, y1) Or ULT(map_.height, y2) Then
            printLog_ pCustomLogger, 313, LogLevel_Warning, "writeMapData", "Bad gateway dimensions (exceed map width or height)"
            ' for now, write anyway
        End If

        ok = BinaryIOStream_writeULE8(pStream, x1)
        If ok Then ok = BinaryIOStream_writeULE8(pStream, y1)
        If ok Then ok = BinaryIOStream_writeULE8(pStream, x2)
        If ok Then ok = BinaryIOStream_writeULE8(pStream, y2)
        If Not ok Then
            printLog_ pCustomLogger, 320, LogLevel_Error, "writeMapData", "Error writing gateway to: %s", path
            Exit Function
        End If
    Next

    Set pStream = Nothing
    writeMapData = True
End Function

'--------------------------------------------------------------------------
' MARK: - Helper functions for loading / saving JSON files

' Left-pads the current string with codepoint ch up to the minimumStringLength
' If the current string length() is already >= minimumStringLength, no padding occurs.
' (It pads on the right: str = str + strPadding.)
Private Function leftPadStrToMinimumLength(str_ As String, ByVal minimumStringLength As Long, Optional ByVal paddingChar As String = "0") As String
    Dim leftPaddingRequired As Long, strPadding As String, i As Long
    If Len(str_) >= minimumStringLength Then
        leftPadStrToMinimumLength = str_
        Exit Function
    End If

    leftPaddingRequired = minimumStringLength - Len(str_)
    For i = 0 To leftPaddingRequired - 1
        strPadding = String$(leftPaddingRequired, paddingChar)     ' strPadding.resize(leftPaddingRequired, paddingChar)
    Next
    str_ = str_ & strPadding
    leftPadStrToMinimumLength = str_
End Function

Private Function numberOfDigitsInNumber(ByVal num As Long) As Long
    Dim digits As Long
    digits = 0
    Do
        num = num \ 10
        digits = digits + 1
    Loop While num <> 0
    numberOfDigitsInNumber = digits
End Function

' items() of a JSON object or array: the i-th key (an array's is its index)
Private Function ItemKey(ByVal container As nlohmann_json, ByVal i As Long) As String
    If container.is_object() Then ItemKey = container.key_at(i) Else ItemKey = CStr(i)
End Function

Private Function jsonGetRootMapObjectsContainer(ByVal filename As String, ByVal rootObject As nlohmann_json, ByVal jsonFileFormat As Long, ByVal pRootContainerName As String, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As nlohmann_json
    Dim mMapObjectsContainer As nlohmann_json
    Set mMapObjectsContainer = rootObject              ' nlohmann::json& mMapObjectsContainer = rootObject
    If ULT(1, jsonFileFormat) Then
        If Not rootObject.contains(pRootContainerName) Then
            ' Missing required "droid" key for list of droids
            printLog_ pCustomLogger, 376, LogLevel_Error, "jsonGetRootMapObjectsContainer", "%s: Missing required ""%s"" key in root object", filename, pRootContainerName
            Exit Function
        End If
        mMapObjectsContainer.assign rootObject.at(pRootContainerName)    ' assigned through the reference: into rootObject
        If Not mMapObjectsContainer.is_array() Then
            printLog_ pCustomLogger, 382, LogLevel_Error, "jsonGetRootMapObjectsContainer", "%s: ""%s"" value should be an array", filename, pRootContainerName
            Exit Function
        End If
    End If
    Set jsonGetRootMapObjectsContainer = mMapObjectsContainer
End Function

Private Function jsonGetFileFormatVersion(ByVal filename As String, ByVal rootObject As nlohmann_json, ByVal pCustomLogger As LoggingProtocol, ByVal maxSupportedFileFormatVersion As Long, result As Long) As Boolean
    Dim fileFormatVersion As Long, it_version As nlohmann_json
    fileFormatVersion = 1
    Set it_version = rootObject.find("version")
    If Not it_version Is Nothing Then
        If Not it_version.is_number() Then
            printLog_ pCustomLogger, 397, LogLevel_Error, "jsonGetFileFormatVersion", "%s: ""version"" key is not a number", filename
            Exit Function
        End If
        fileFormatVersion = it_version.get_uint()
        If fileFormatVersion = 1 Then
            printLog_ pCustomLogger, 403, LogLevel_Error, "jsonGetFileFormatVersion", "%s: Unsupported file ""version"" (%u) - version 1 lacks a ""version"" key", filename, fileFormatVersion
            Exit Function
        End If
        If fileFormatVersion = 0 Or ULT(maxSupportedFileFormatVersion, fileFormatVersion) Then
            printLog_ pCustomLogger, 408, LogLevel_Error, "jsonGetFileFormatVersion", "%s: Unsupported file ""version"" (%u)", filename, fileFormatVersion
            Exit Function
        End If
    End If
    result = fileFormatVersion
    jsonGetFileFormatVersion = True
End Function

' try { v = it->get<T>() } (idx < 0), or it->at(idx).get<T>(): false with e.what() when it throws
Private Function TryGetT(ByVal it As nlohmann_json, ByVal idx As Long, ByVal T As Long, v As Long, what As String) As Boolean
    Dim j As nlohmann_json
    On Error GoTo catch_
    If idx < 0 Then Set j = it Else Set j = it.at_idx(idx)
    v = j.get_uint()
    If T = 16 Then v = v And &HFFFF&
    TryGetT = True
    Exit Function
catch_:
    If Err.Source <> "nlohmann::json" Then Err.Raise Err.Number, Err.Source, Err.Description
    what = Err.Description
End Function

Private Sub PushLong(arr() As Long, count As Long, ByVal v As Long)
    If count = 0 Then
        ReDim arr(0 To 3)
    ElseIf count > UBound(arr) Then
        ReDim Preserve arr(0 To 2 * count - 1)
    End If
    arr(count) = v
    count = count + 1
End Sub

' T: 32 (uint32_t) or 16 (uint16_t); the list in result(0 .. resultCount - 1)
Private Function jsonGetListOfType(ByVal obj As nlohmann_json, ByVal key As String, ByVal minItems As Long, ByVal maxItems As Long, ByVal ctxFilename As String, ByVal ctxJsonPath As String, ByVal pCustomLogger As LoggingProtocol, ByVal allowNonList As Boolean, ByVal T As Long, result() As Long, resultCount As Long) As Boolean
    Dim it As nlohmann_json, idx As Long, v As Long, what As String
    resultCount = 0
    Set it = obj.find(key)
    If it Is Nothing Then
        Exit Function
    End If

    If it.size() < minItems Then
        ' Invalid value - requires at least minItems components
        printLog_ pCustomLogger, 428, LogLevel_Error, "jsonGetListOfType", "%s: Invalid ""%s"" (requires at least %zu members) for: %s", ctxFilename, key, minItems, ctxJsonPath
        Exit Function
    End If

    If Not it.is_array() Then
        If Not allowNonList Then
            ' Invalid value - expecting an array
            printLog_ pCustomLogger, 437, LogLevel_Error, "jsonGetListOfType", "%s: Invalid ""%s"" (expecting an array) for: %s", ctxFilename, key, ctxJsonPath
            Exit Function
        End If
        ' process non-array
        If Not TryGetT(it, -1, T, v, what) Then
            printLog_ pCustomLogger, 445, LogLevel_Error, "jsonGetListOfType", "%s: Invalid ""%s"" (unable to convert to desired output type) for: %s; error: %s", ctxFilename, key, ctxJsonPath, what
            Exit Function
        End If
        PushLong result, resultCount, v
        jsonGetListOfType = True
        Exit Function
    End If

    For idx = 0 To minItems - 1
        If Not TryGetT(it, idx, T, v, what) Then
            printLog_ pCustomLogger, 458, LogLevel_Error, "jsonGetListOfType", "%s: Invalid ""%s"" (unable to convert to desired output type) for: %s; error: %s", ctxFilename, key, ctxJsonPath, what
            Exit Function
        End If
        PushLong result, resultCount, v
    Next

    If it.size() > maxItems Then
        ' Too many items
        printLog_ pCustomLogger, 465, LogLevel_Error, "jsonGetListOfType", "%s: Invalid ""%s"" (too many array members) for: %s", ctxFilename, key, ctxJsonPath
    ElseIf it.size() > minItems Then
        ' check for non-zero additional members, and warn that they are being ignored
        For idx = minItems To it.size() - 1
            If Not TryGetT(it, idx, T, v, what) Then
                printLog_ pCustomLogger, 480, LogLevel_Warning, "jsonGetListOfType", "%s: Invalid ignored ""%s"" member (unable to convert to desired output type) for: %s; error: %s", ctxFilename, key, ctxJsonPath, what
                Exit For
            End If
            If v <> 0 Then
                printLog_ pCustomLogger, 475, LogLevel_Info_Verbose, "jsonGetListOfType", "%s: Ignoring non-0 ""%s[%zu]"" for: %s", ctxFilename, key, idx, ctxJsonPath
            End If
        Next
    End If
    jsonGetListOfType = True
End Function

' the player is extracted from either "player" or "startpos"
' for multiplayer / skirmish maps, "startpos" should be set to the player slot index
' for campaign maps, "player" can be set to the player number
' in all cases, "player" can be set to the string "scavenger" to identify the object as belonging to the scavengers player
Private Function jsonGetPlayerFromObj(ByVal obj As nlohmann_json, ByVal mapType_ As MapType, ByVal ctxFilename As String, ByVal ctxJsonPath As String, ByVal pCustomLogger As LoggingProtocol, ByVal playerFieldOnly As Boolean, player As Long) As Boolean
    Dim it As nlohmann_json, isScavenger As Boolean
    Set it = obj.find("player")
    If Not it Is Nothing Then
        If obj.contains("startpos") Then
            printLog_ pCustomLogger, 497, LogLevel_Warning, "jsonGetPlayerFromObj", "%s: Processing ""player"", ignoring ""startpos"" for: %s", ctxFilename, ctxJsonPath
            ' continue to process "player"
        End If
        If it.is_string() Then isScavenger = (Left$(it.get_string(), 9) = "scavenger")
        If isScavenger Then
            player = PLAYER_SCAVENGERS
            jsonGetPlayerFromObj = True
            Exit Function
        ElseIf it.is_number() Then
            If mapType_ <> MapType_CAMPAIGN And mapType_ <> MapType_SAVEGAME Then
                ' "player" is generally expected only for campaign and savegames (and for specifying scavengers)
                printLog_ pCustomLogger, 509, LogLevel_Info_Verbose, "jsonGetPlayerFromObj", "Found ""player"" for non-campaign/savegame."
                ' however, proceed anyway (for now)
            End If
            player = ToInt8(it.get_int())
            jsonGetPlayerFromObj = True
            Exit Function
        End If
        ' otherwise, invalid json type
        Exit Function
    End If
    If playerFieldOnly Then
        Exit Function
    End If
    Set it = obj.find("startpos")
    If Not it Is Nothing Then
        If Not it.is_number() Then
            ' invalid json type
            Exit Function
        End If
        player = ToInt8(it.get_int())
        jsonGetPlayerFromObj = True
        Exit Function
    End If
    ' No player info found!
End Function

Private Function jsonSetPlayerOnObject(ByVal jsonObj As nlohmann_json, ByVal mapType_ As MapType, ByVal player As Long) As Boolean
    If player = PLAYER_SCAVENGERS Then
        jsonObj.set_item_ref "player", json_value("scavenger", True)
    Else
        If player < 0 Then
            Exit Function
        End If
        Select Case mapType_
        Case MapType_CAMPAIGN, MapType_SAVEGAME
            jsonObj.set_item_ref "player", json_value(player, True)
        Case MapType_SKIRMISH
            jsonObj.set_item_ref "startpos", json_value(player, True)
        End Select
    End If

    jsonSetPlayerOnObject = True
End Function

' template<typename T> (Structure, Droid, Feature)
Private Function jsonGetBaseMapObjectInfo(ByVal mapObj As Object, ByVal jsonFileVersion As Long, ByVal jsonObj As nlohmann_json, ByVal ctxFilename As String, ByVal ctxJsonPath As String, ByVal pCustomLogger As LoggingProtocol, ByVal pNameKey As String) As Boolean
    Dim maxComponentsPosition As Long, maxComponentsRotation As Long, rotationAllowNonArray As Boolean
    Dim it_name As nlohmann_json, it_id As nlohmann_json, id As Long
    Dim position() As Long, nPosition As Long, rotation() As Long, nRotation As Long
    maxComponentsPosition = 3
    maxComponentsRotation = 3
    rotationAllowNonArray = False
    If ULT(1, jsonFileVersion) Then
        maxComponentsPosition = 2 ' [x, y] only
        maxComponentsRotation = 1 ' [x] only
        rotationAllowNonArray = True
    End If
    Set it_name = jsonObj.find(pNameKey)
    If it_name Is Nothing Then
        ' Missing required "template" key for map object
        printLog_ pCustomLogger, 579, LogLevel_Error, "jsonGetBaseMapObjectInfo", "%s: Missing required ""%s"" key for: %s", ctxFilename, pNameKey, ctxJsonPath
        Exit Function
    End If
    If Not it_name.is_string() Then
        ' Expecting a string value for name
        printLog_ pCustomLogger, 585, LogLevel_Error, "jsonGetBaseMapObjectInfo", "%s: Unexpected type of required ""%s"" key (expecting string) for: %s", ctxFilename, pNameKey, ctxJsonPath
        Exit Function
    End If
    mapObj.name_ = it_name.get_string()
    ' "id" is explicitly optional - synchronized unit ids will be generated if it is omitted
    Set it_id = jsonObj.find("id")
    If Not it_id Is Nothing Then
        If Not it_id.is_number() Then
            ' Expecting a numeric value for id
            printLog_ pCustomLogger, 596, LogLevel_Error, "jsonGetBaseMapObjectInfo", "%s: Unexpected type of ""id"" key (expecting number) for: %s", ctxFilename, ctxJsonPath
            Exit Function
        End If
        id = it_id.get_uint()
        If id <> 0 Then
            mapObj.id = id
            mapObj.hasId = True
        Else
            printLog_ pCustomLogger, 606, LogLevel_Warning, "jsonGetBaseMapObjectInfo", "%s: Invalid ""id"" = 0 for: %s - ignoring id value", ctxFilename, ctxJsonPath
        End If
    End If
    ' "position" must contain at least two components [x, y]
    If jsonGetListOfType(jsonObj, "position", 2, maxComponentsPosition, ctxFilename, ctxJsonPath, pCustomLogger, False, 32, position, nPosition) Then
        mapObj.position_x = position(0)
        mapObj.position_y = position(1)
    Else
        ' Missing required "position" key for map object
        printLog_ pCustomLogger, 619, LogLevel_Error, "jsonGetBaseMapObjectInfo", "%s: Missing required ""position"" key for: %s", ctxFilename, ctxJsonPath
        Exit Function
    End If
    ' "rotation" must contain at least one components [x]
    If jsonGetListOfType(jsonObj, "rotation", 1, maxComponentsRotation, ctxFilename, ctxJsonPath, pCustomLogger, rotationAllowNonArray, 16, rotation, nRotation) Then
        mapObj.direction = rotation(0)
    End If

    jsonGetBaseMapObjectInfo = True
End Function

' template<typename T> (Structure, Droid, Feature)
Private Sub jsonSetBaseMapObjectInfo(ByVal jsonObj As nlohmann_json, ByVal jsonFileVersion As Long, ByVal mapObj As Object, ByVal pNameKey As String)
    Dim position As nlohmann_json, rotation As nlohmann_json
    jsonObj.set_item_ref pNameKey, json_value(mapObj.name_, True)
    If mapObj.hasId Then
        jsonObj.set_item_ref "id", json_value_unsigned(mapObj.id, True)
    End If
    ' "position" must contain at least two components [x, y] - jsonFileVersion_v1 always expects 3 (and ignores the third)
    Set position = json_array(True)
    position.push_back_ref json_value_unsigned(mapObj.position_x, True)
    position.push_back_ref json_value_unsigned(mapObj.position_y, True)
    If jsonFileVersion = 1 Then
        position.push_back_ref json_value(0&, True)
    End If
    jsonObj.set_item_ref "position", position
    Set position = json_value(Empty, True)              ' std::move: position is left null
    If jsonFileVersion = 1 Then
        ' "rotation" must contain at least one components [x] - jsonFileVersion_v1 always expects 3 (and ignores all but the first)
        Set rotation = json_array(True)
        rotation.push_back_ref json_value_unsigned(mapObj.direction, True)
        position.push_back_ref json_value(0&, True)
        position.push_back_ref json_value(0&, True)
    Else
        ' "rotation" - v2+ JSON file format supports direct (non-array) single value
        Set rotation = json_value_unsigned(mapObj.direction, True)
    End If
    jsonObj.set_item_ref "rotation", rotation
End Sub

'--------------------------------------------------------------------------
' MARK: - Helper functions for loading / saving binary (BJO) files

Private Function bjoScavengerSlot(ByVal mapMaxPlayer As Long) As Long
    ' For old binary file formats:
    ' Scavengers used to always be in position 7, when scavengers were only supported in less than 8 player maps.
    ' Scavengers should be in position N in N-player maps, where N >= 8.
    If ULT(mapMaxPlayer, 7) Then bjoScavengerSlot = 7 Else bjoScavengerSlot = mapMaxPlayer
End Function

Private Function bjoConvertPlayer(ByVal bjoPlayer As Long, ByVal mapMaxPlayers As Long) As Long
    If bjoScavengerSlot(mapMaxPlayers) = bjoPlayer Then
        bjoConvertPlayer = PLAYER_SCAVENGERS
        Exit Function
    End If

    bjoConvertPlayer = ToInt8(bjoPlayer)
End Function

Private Function convertPlayerToBJOPlayer(ByVal player As Long, ByVal mapMaxPlayers As Long) As Long
    If player = PLAYER_SCAVENGERS Then
        convertPlayerToBJOPlayer = bjoScavengerSlot(mapMaxPlayers)
        Exit Function
    End If

    convertPlayerToBJOPlayer = player
End Function

' template<typename T> (Structure, Droid, Feature)
Private Function getLargestSpecifiedId(ByVal container As ObjectVector) As Long
    Dim largestId As Long, i As Long, item As Object
    largestId = 0
    For i = 0 To container.size() - 1
        Set item = container.at(i)
        If item.hasId Then
            If ULT(largestId, item.id) Then largestId = item.id
        End If
    Next
    getLargestSpecifiedId = largestId
End Function

' the name field of a .bjo record: name, truncated to nameFieldLength bytes and padded with zeros
Private Function writeBJOName(ByVal pStream As BinaryIOStream, nameBytes() As Byte, ByVal nameStrWriteLen As Long, ByVal nameFieldLength As Long, bigZeroBuff() As Byte, nameOk As Boolean) As Boolean
    Dim n As Long
    nameOk = pStream.writeBytes(nameBytes, 0, nameStrWriteLen, n)
    If nameOk Then nameOk = (n = nameStrWriteLen)
    If Not nameOk Then Exit Function
    If nameStrWriteLen < nameFieldLength Then
        If Not pStream.writeBytes(bigZeroBuff, 0, nameFieldLength - nameStrWriteLen, n) Then Exit Function
        If n <> nameFieldLength - nameStrWriteLen Then Exit Function
    End If
    writeBJOName = True
End Function

'--------------------------------------------------------------------------
' MARK: - Structure loading functions

Public Function loadBJOStructureInit(ByVal filename As String, ByVal mapMaxPlayers As Long, ByVal mapIO As IOProvider, ByVal pCustomLogger As LoggingProtocol, resultObjects As ObjectVector, resultFileFormatVersion As Long) As Boolean
    Dim path As String, pStream As BinaryIOStream, objects As ObjectVector
    Dim aFileType(0 To 3) As Byte, n As Long, ok As Boolean, version As Long, quantity As Long
    Dim nameLength As Long, nameBuff() As Byte, i As Long, k As Long
    Dim id As Long, x As Long, y As Long, z As Long, direction As Long, player As Long
    Dim inFire As Long, periodicalDamageStart As Long, periodicalDamage As Long, status As Long, capacity As Long
    Dim dummy As Long, dummy8 As Long, dummyS32 As Long, structure_ As Structure
    Set objects = New ObjectVector
    path = filename

    Set pStream = mapIO.openBinaryStream(filename, OpenMode_READ)
    If pStream Is Nothing Then
        ' Failed to open file
        Exit Function
    End If

    printLog_ pCustomLogger, 727, LogLevel_Info, "loadBJOStructureInit", "Loading: %s", path

    version = 0
    quantity = 0
    ok = pStream.readBytes(aFileType, 0, 4, n)
    If ok Then ok = (n = 4)
    If ok Then ok = (aFileType(0) = 115 And aFileType(1) = 116 And aFileType(2) = 114 And aFileType(3) = 117)   ' "stru"
    If ok Then ok = BinaryIOStream_readULE32(pStream, version)
    If ok Then ok = BinaryIOStream_readULE32(pStream, quantity)
    If Not ok Then
        printLog_ pCustomLogger, 740, LogLevel_Error, "loadBJOStructureInit", "%s: Bad header", path
        Exit Function
    End If
    resultFileFormatVersion = version

    If ULT(version, 7) Or ULT(8, version) Then
        printLog_ pCustomLogger, 747, LogLevel_Error, "loadBJOStructureInit", "%s: Unsupported binary struct file version: %u", path, version
        Exit Function
    End If

    nameLength = 60
    If Not ULT(19, version) Then
        nameLength = 40
    End If
    ReDim nameBuff(0 To nameLength - 1)

    i = 0
    Do While ULT(i, quantity)
        id = 0
        x = 0: y = 0: z = 0
        direction = 0
        player = 0

        inFire = 0
        periodicalDamageStart = 0
        periodicalDamage = 0
        status = 0
        capacity = 0

        ok = pStream.readBytes(nameBuff, 0, nameLength, n)
        If ok Then ok = (n = nameLength)
        If ok Then ok = BinaryIOStream_readULE32(pStream, id)
        If ok Then ok = BinaryIOStream_readULE32(pStream, x)
        If ok Then ok = BinaryIOStream_readULE32(pStream, y)
        If ok Then ok = BinaryIOStream_readULE32(pStream, z)
        If ok Then ok = BinaryIOStream_readULE32(pStream, direction)
        If ok Then ok = BinaryIOStream_readULE32(pStream, player)
        If ok Then ok = BinaryIOStream_readSLE32(pStream, inFire)                  ' BOOL inFire
        If ok Then ok = BinaryIOStream_readULE32(pStream, periodicalDamageStart)   ' burnStart
        If ok Then ok = BinaryIOStream_readULE32(pStream, periodicalDamage)        ' burnDamage
        If ok Then ok = BinaryIOStream_readULE8(pStream, status)                   ' status - causes structure padding
        If ok Then ok = BinaryIOStream_readULE8(pStream, dummy8)                   ' structure padding
        If ok Then ok = BinaryIOStream_readULE8(pStream, dummy8)                   ' structure padding
        If ok Then ok = BinaryIOStream_readULE8(pStream, dummy8)                   ' structure padding
        If ok Then ok = BinaryIOStream_readSLE32(pStream, dummyS32)                ' currentBuildPts - aligned on 4 byte boundary
        For k = 1 To 7                                                             ' body, armour, resistance, dummy1, subjectInc, timeStarted, output
            If ok Then ok = BinaryIOStream_readULE32(pStream, dummy)
        Next
        If ok Then ok = BinaryIOStream_readULE32(pStream, capacity)                ' capacity
        If ok Then ok = BinaryIOStream_readULE32(pStream, dummy)                   ' quantity
        If Not ok Then
            printLog_ pCustomLogger, 798, LogLevel_Error, "loadBJOStructureInit", "%s: Failed to read structure %u", path, i
            Exit Function
        End If

        Set structure_ = New Structure
        If id <> 0 Then
            structure_.id = id
            structure_.hasId = True
        Else
            printLog_ pCustomLogger, 809, LogLevel_Warning, "loadBJOStructureInit", "%s: Structure %u has id = 0 - ignoring id value", path, i
        End If
        structure_.name_ = CStringOf(nameBuff, nameLength)
        ' TODO: Possibly handle collecting modules?
        structure_.position_x = x
        structure_.position_y = y
        ' ignore z component
        structure_.direction = DEG_u16(direction)
        structure_.player = bjoConvertPlayer(player, mapMaxPlayers)
        ' check inFire
        If inFire <> 0 Then
            printLog_ pCustomLogger, 821, LogLevel_Warning, "loadBJOStructureInit", "%s: Ignoring inFire(%i) for struct %u", path, inFire, i
        End If
        ' check periodicalDamageStart
        If periodicalDamageStart <> 0 Then
            printLog_ pCustomLogger, 826, LogLevel_Warning, "loadBJOStructureInit", "%s: Ignoring periodicalDamageStart(%u) for struct %u", path, periodicalDamageStart, i
        End If
        ' check periodicalDamage
        If periodicalDamage <> 0 Then
            printLog_ pCustomLogger, 831, LogLevel_Warning, "loadBJOStructureInit", "%s: Ignoring periodicalDamage(%u) for struct %u", path, periodicalDamage, i
        End If
        If status <> SS_BUILT Then
            printLog_ pCustomLogger, 835, LogLevel_Warning, "loadBJOStructureInit", "%s: Ignoring status(%u) for struct %u", path, status, i
        End If
        If capacity <> 0 Then
            printLog_ pCustomLogger, 839, LogLevel_Warning, "loadBJOStructureInit", "%s: Ignoring capacity(%u) for struct %u", path, capacity, i
        End If
        ' TODO: Sanity check struct position ?
        objects.push_back structure_
        i = i + 1
    Loop
    ' Check: extra bytes at end
    If Not pStream.endOfStream() Then
        printLog_ pCustomLogger, 847, LogLevel_Warning, "loadBJOStructureInit", "%s: Unexpectedly did not reach end of stream - data may be corrupted", path
    End If
    Set resultObjects = objects
    loadBJOStructureInit = True
End Function

Private Function isKnownStructureJSONKey(ByVal key As String) As Boolean
    Select Case key
    Case "name", "id", "position", "rotation", "player", "startpos", "modules"
        isKnownStructureJSONKey = True
    End Select
End Function

Public Function loadJsonStructureInit(ByVal filename As String, ByVal mapType_ As MapType, ByVal mapIO As IOProvider, ByVal pCustomLogger As LoggingProtocol, resultObjects As ObjectVector, resultFileFormatVersion As Long) As Boolean
    Dim path As String, objects As ObjectVector, mRoot As nlohmann_json, detectedFormatVersion As Long
    Dim pStructuresRoot As nlohmann_json, structureJson As nlohmann_json, it_modules As nlohmann_json
    Dim i As Long, j As Long, key As String, structure_ As Structure, player As Long, modulesCount As Double
    Set objects = New ObjectVector
    path = filename

    Set mRoot = loadJsonObjectFromFile(filename, mapIO, pCustomLogger)
    If mRoot Is Nothing Then
        ' Failed to load JSON - rely on loadJsonFromFile to handle output of errors
        Exit Function
    End If

    printLog_ pCustomLogger, 866, LogLevel_Info, "loadJsonStructureInit", "Loading: %s", path

    If Not jsonGetFileFormatVersion(filename, mRoot, pCustomLogger, 2, detectedFormatVersion) Then
        ' jsonGetFileFormatVersion should have already logged the reason
        Exit Function
    End If
    resultFileFormatVersion = detectedFormatVersion

    Set pStructuresRoot = jsonGetRootMapObjectsContainer(filename, mRoot, detectedFormatVersion, "structures")
    If pStructuresRoot Is Nothing Then
        ' jsonGetRootMapObjectsContainer should have already logged the reason
        Exit Function
    End If

    For i = 0 To pStructuresRoot.size() - 1
        key = ItemKey(pStructuresRoot, i)
        Set structureJson = pStructuresRoot.value_at(i)
        If Not structureJson.is_object() Then
            GoTo continue_
        End If
        Set structure_ = New Structure
        If Not jsonGetBaseMapObjectInfo(structure_, detectedFormatVersion, structureJson, path, key, pCustomLogger, "name") Then
            ' jsonGetBaseMapObjectInfo should have handled logging an error
            GoTo continue_
        End If
        ' the player is extracted from either "player" or "startpos" - see jsonGetPlayerFromObj
        If jsonGetPlayerFromObj(structureJson, mapType_, path, key, pCustomLogger, False, player) Then
            structure_.player = player
        Else
            ' Missing required "player" or "startpos" key for droid
            printLog_ pCustomLogger, 907, LogLevel_Error, "loadJsonStructureInit", "%s: Missing required player/startpos key for structure: %s", path, key
            GoTo continue_
        End If
        ' "modules" (capacity)
        Set it_modules = structureJson.find("modules")
        If Not it_modules Is Nothing Then
            If Not it_modules.is_number() Then
                printLog_ pCustomLogger, 916, LogLevel_Error, "loadJsonStructureInit", "%s: Unexpected type of ""modules"" key (expecting number) for: %s", path, key
                GoTo continue_
            End If
            modulesCount = it_modules.get_uint64()
            If modulesCount > 255 Or modulesCount < 0 Then      ' (a negative double converts to a huge uint64_t)
                ' "modules" value exceeds maximum allowable value
                printLog_ pCustomLogger, 923, LogLevel_Error, "loadJsonStructureInit", "%s: ""modules"" value exceeds maximum allowable value: %s", path, key
                GoTo continue_
            End If
            structure_.modules = CLng(modulesCount)
        End If

        ' Sanity check for unknown / unprocessed keys
        For j = 0 To structureJson.size() - 1
            If Not isKnownStructureJSONKey(structureJson.key_at(j)) Then
                printLog_ pCustomLogger, 934, LogLevel_Warning, "loadJsonStructureInit", "%s: Unexpected structure key ""%s"" for structure: %s", path, structureJson.key_at(j), key
            End If
        Next

        objects.push_back structure_
continue_:
    Next

    Set resultObjects = objects
    loadJsonStructureInit = True
End Function

'--------------------------------------------------------------------------
' MARK: - Structure writing functions

Public Function writeMapStructureInitBJO(ByVal structures As ObjectVector, ByVal mapMaxPlayers As Long, ByVal filename As String, ByVal mapIO As IOProvider, ByVal bjoFileVersion As Long, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim path As String, pStream As BinaryIOStream, aFileType(0 To 3) As Byte, n As Long, ok As Boolean, quantity As Long
    Dim nameFieldLength As Long, bigZeroBuff(0 To 255) As Byte, nextAvailableStructureId As Long, i As Long, k As Long
    Dim structure_ As Structure, nameBytes() As Byte, nameSize As Long, nameStrWriteLen As Long, nameOk As Boolean
    Dim structureId As Long, bjoPlayer As Long
    path = filename

    If ULT(bjoFileVersion, 7) Or ULT(8, bjoFileVersion) Then
        ' Currently, we only support versions 7 + 8
        printLog_ pCustomLogger, 953, LogLevel_Error, "writeMapStructureInitBJO", "Unsupported binary struct file version: %u; cannot write to: %s", bjoFileVersion, path
        Exit Function
    End If

    Set pStream = mapIO.openBinaryStream(filename, OpenMode_WRITE)

    If pStream Is Nothing Then
        printLog_ pCustomLogger, 961, LogLevel_Error, "writeMapStructureInitBJO", "Failed to open file for writing: %s", path
        Exit Function
    End If

    printLog_ pCustomLogger, 965, LogLevel_Info, "writeMapStructureInitBJO", "Writing: %s", path

    ' write header
    aFileType(0) = 115  ' s
    aFileType(1) = 116  ' t
    aFileType(2) = 114  ' r
    aFileType(3) = 117  ' u
    quantity = structures.size()
    ok = pStream.writeBytes(aFileType, 0, 4, n)
    If ok Then ok = (n = 4)
    If ok Then ok = BinaryIOStream_writeULE32(pStream, bjoFileVersion)
    If ok Then ok = BinaryIOStream_writeULE32(pStream, quantity)
    If Not ok Then
        printLog_ pCustomLogger, 978, LogLevel_Error, "writeMapStructureInitBJO", "Failed writing header to: %s", path
        Exit Function
    End If

    nameFieldLength = 60
    If Not ULT(19, bjoFileVersion) Then
        nameFieldLength = 40
    End If

    nextAvailableStructureId = U32Inc(getLargestSpecifiedId(structures))

    For i = 0 To quantity - 1
        Set structure_ = structures.at(i)
        nameSize = utf8_encode(structure_.name_, nameBytes)
        nameStrWriteLen = nameFieldLength
        If nameSize < nameStrWriteLen Then nameStrWriteLen = nameSize
        If nameStrWriteLen < nameSize Then
            ' Warn about truncation!!
            printLog_ pCustomLogger, 998, LogLevel_Warning, "writeMapStructureInitBJO", "Structure's name exceeds the length supported by binary struct file version: %u; output will be truncated", bjoFileVersion
        End If
        If Not writeBJOName(pStream, nameBytes, nameStrWriteLen, nameFieldLength, bigZeroBuff, nameOk) Then
            If Not nameOk Then
                ' Failed to write
                printLog_ pCustomLogger, 1003, LogLevel_Error, "writeMapStructureInitBJO", "Error writing structure %u's name to: %s", i, path
            Else
                ' Failed to pad with zeros
                printLog_ pCustomLogger, 1010, LogLevel_Error, "writeMapStructureInitBJO", "Error writing structure %u to: %s", i, path
            End If
            Exit Function
        End If

        structureId = 0
        If structure_.hasId Then
            structureId = structure_.id
        Else
            structureId = nextAvailableStructureId
            nextAvailableStructureId = U32Inc(nextAvailableStructureId)
        End If

        bjoPlayer = convertPlayerToBJOPlayer(structure_.player, mapMaxPlayers)
        ' TODO: Do campaign maps require scanvengers at player ID 7? (Theoretically only an issue for old campaign maps with > 7 players? And there shouldn't be any of those?)

        ' TODO: Better way to handle if structure.modules > 0 ?
        If structure_.modules > 0 Then
            printLog_ pCustomLogger, 1030, LogLevel_Warning, "writeMapStructureInitBJO", "Structure modules (%u) > 0. Conversion to old binary struct format (version: %u) may be missing the modules.", structure_.modules, bjoFileVersion
        End If

        ok = BinaryIOStream_writeULE32(pStream, structureId)
        If ok Then ok = BinaryIOStream_writeULE32(pStream, structure_.position_x)
        If ok Then ok = BinaryIOStream_writeULE32(pStream, structure_.position_y)
        If ok Then ok = BinaryIOStream_writeULE32(pStream, 0)
        If ok Then ok = BinaryIOStream_writeULE32(pStream, structure_.direction)   ' TODO: Conversion / manipulation needed?
        If ok Then ok = BinaryIOStream_writeULE32(pStream, bjoPlayer)
        If ok Then ok = BinaryIOStream_writeSLE32(pStream, 0)       ' BOOL inFire
        If ok Then ok = BinaryIOStream_writeULE32(pStream, 0)       ' burnStart
        If ok Then ok = BinaryIOStream_writeULE32(pStream, 0)       ' burnDamage
        If ok Then ok = BinaryIOStream_writeULE8(pStream, SS_BUILT) ' status - causes structure padding
        ' the following magic padding bytes match the behavior of flaME's structure output - they don't seem to be used, but why not?
        If ok Then ok = BinaryIOStream_writeULE8(pStream, 26)       ' structure padding
        If ok Then ok = BinaryIOStream_writeULE8(pStream, 127)      ' structure padding
        If ok Then ok = BinaryIOStream_writeULE8(pStream, 0)        ' structure padding
        If ok Then ok = BinaryIOStream_writeSLE32(pStream, 0)       ' currentBuildPts - aligned on 4 byte boundary
        For k = 1 To 9                                              ' body, armour, resistance, dummy1, subjectInc, timeStarted, output, capacity, quantity
            If ok Then ok = BinaryIOStream_writeULE32(pStream, 0)
        Next
        If Not ok Then
            printLog_ pCustomLogger, 1057, LogLevel_Error, "writeMapStructureInitBJO", "Error writing structure %u to: %s", i, path
            Exit Function
        End If
    Next

    Set pStream = Nothing
    writeMapStructureInitBJO = True
End Function

Public Function writeMapStructureInitJSON(ByVal structures As ObjectVector, ByVal mapType_ As MapType, ByVal filename As String, ByVal mapIO As IOProvider, ByVal jsonFileVersion As Long, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim path As String, mRoot As nlohmann_json, pStructuresRoot As nlohmann_json, structureObj As nlohmann_json
    Dim counter As Long, structureCountMinLength As Long, i As Long, structure_ As Structure, counterStr As String
    path = filename
    If ULT(2, jsonFileVersion) Then
        ' Currently, we only support up to version 2 JSON files
        printLog_ pCustomLogger, 1072, LogLevel_Error, "writeMapStructureInitJSON", "Unsupported JSON structure file version: %u; cannot write to: %s", jsonFileVersion, path
        Exit Function
    End If

    Set mRoot = json_object(True)
    Set pStructuresRoot = mRoot
    If ULT(1, jsonFileVersion) Then
        ' Version 2+ JSON format contains a "version" key
        ' And stores the objects in an array
        mRoot.set_item_ref "version", json_value_unsigned(jsonFileVersion, True)
        mRoot.set_item_ref "structures", json_array(True)
        Set pStructuresRoot = mRoot.at("structures")
    End If

    counter = 0
    structureCountMinLength = numberOfDigitsInNumber(structures.size())
    If structureCountMinLength < 1 Then structureCountMinLength = 1
    For i = 0 To structures.size() - 1
        Set structure_ = structures.at(i)
        Set structureObj = json_object(True)
        ' write base object info (name, <id>, position, rotation)
        jsonSetBaseMapObjectInfo structureObj, jsonFileVersion, structure_, "name"
        ' player
        If Not jsonSetPlayerOnObject(structureObj, mapType_, structure_.player) Then
            printLog_ pCustomLogger, 1097, LogLevel_Error, "writeMapStructureInitJSON", "Invalid player number (%i)", structure_.player
            Exit Function
        End If
        ' "modules" (capacity)
        If structure_.modules > 0 Then
            structureObj.set_item_ref "modules", json_value_unsigned(structure_.modules, True)
        End If

        If pStructuresRoot.is_object() Then
            ' add to the root object
            counterStr = CStr(counter)
            leftPadStrToMinimumLength counterStr, structureCountMinLength, "0"
            pStructuresRoot.set_item_ref "structure_" & counterStr, structureObj
        Else 'if (pStructuresRoot->is_array())
            pStructuresRoot.push_back_ref structureObj
        End If
        counter = counter + 1
    Next

    ' write out to file
    writeMapStructureInitJSON = saveOrderedJsonObjectToFile(mRoot, filename, mapIO)
End Function

Public Function writeMapStructureInit(ByVal structures As ObjectVector, ByVal mapMaxPlayers As Long, ByVal mapType_ As MapType, ByVal mapFolderPath As String, ByVal mapIO As IOProvider, ByVal format As OutputFormat, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim retVal As Boolean
    retVal = False
    Select Case format
    Case OutputFormat_VER1_BINARY_OLD
        retVal = writeMapStructureInitBJO(structures, mapMaxPlayers, IOProvider_pathJoin(mapIO, mapFolderPath, "struct.bjo"), mapIO, 8, pCustomLogger)
    Case OutputFormat_VER2
        retVal = writeMapStructureInitJSON(structures, mapType_, IOProvider_pathJoin(mapIO, mapFolderPath, "struct.json"), mapIO, 1, pCustomLogger)
    Case OutputFormat_VER3
        retVal = writeMapStructureInitJSON(structures, mapType_, IOProvider_pathJoin(mapIO, mapFolderPath, "struct.json"), mapIO, 2, pCustomLogger)
    End Select
    writeMapStructureInit = retVal
End Function

'--------------------------------------------------------------------------
' MARK: - Droid loading functions

Public Function loadBJODroidInit(ByVal filename As String, ByVal mapMaxPlayers As Long, ByVal mapIO As IOProvider, ByVal pCustomLogger As LoggingProtocol, resultObjects As ObjectVector, resultFileFormatVersion As Long) As Boolean
    Dim path As String, pStream As BinaryIOStream, objects As ObjectVector
    Dim aFileType(0 To 3) As Byte, n As Long, ok As Boolean, version As Long, quantity As Long
    Dim nameLength As Long, nameBuff() As Byte, i As Long
    Dim id As Long, x As Long, y As Long, z As Long, direction As Long, player As Long
    Dim inFire As Long, periodicalDamageStart As Long, periodicalDamage As Long, droid_ As Droid
    Set objects = New ObjectVector
    path = filename

    Set pStream = mapIO.openBinaryStream(filename, OpenMode_READ)
    If pStream Is Nothing Then
        ' Failed to open file
        Exit Function
    End If

    printLog_ pCustomLogger, 1156, LogLevel_Info, "loadBJODroidInit", "Loading: %s", path

    version = 0
    quantity = 0
    ok = pStream.readBytes(aFileType, 0, 4, n)
    If ok Then ok = (n = 4)
    If ok Then ok = (aFileType(0) = 100 And aFileType(1) = 105 And aFileType(2) = 110 And aFileType(3) = 116)  ' "dint"
    If ok Then ok = BinaryIOStream_readULE32(pStream, version)
    If ok Then ok = BinaryIOStream_readULE32(pStream, quantity)
    If Not ok Then
        printLog_ pCustomLogger, 1169, LogLevel_Error, "loadBJODroidInit", "%s: Bad header", path
        Exit Function
    End If
    resultFileFormatVersion = version

    nameLength = 60
    If Not ULT(19, version) Then
        nameLength = 40
    End If
    ReDim nameBuff(0 To nameLength - 1)

    i = 0
    Do While ULT(i, quantity)
        id = 0
        x = 0: y = 0: z = 0
        direction = 0
        player = 0
        inFire = 0
        periodicalDamageStart = 0
        periodicalDamage = 0
        ok = pStream.readBytes(nameBuff, 0, nameLength, n)
        If ok Then ok = (n = nameLength)
        If ok Then ok = BinaryIOStream_readULE32(pStream, id)
        If ok Then ok = BinaryIOStream_readULE32(pStream, x)
        If ok Then ok = BinaryIOStream_readULE32(pStream, y)
        If ok Then ok = BinaryIOStream_readULE32(pStream, z)
        If ok Then ok = BinaryIOStream_readULE32(pStream, direction)
        If ok Then ok = BinaryIOStream_readULE32(pStream, player)
        If ok Then ok = BinaryIOStream_readSLE32(pStream, inFire)                  ' BOOL inFire
        If ok Then ok = BinaryIOStream_readULE32(pStream, periodicalDamageStart)   ' burnStart
        If ok Then ok = BinaryIOStream_readULE32(pStream, periodicalDamage)        ' burnDamage
        If Not ok Then
            printLog_ pCustomLogger, 1199, LogLevel_Error, "loadBJODroidInit", "%s: Failed to read droid %u", path, i
            Exit Function
        End If
        Set droid_ = New Droid
        If id <> 0 Then
            droid_.id = id
            droid_.hasId = True
        Else
            printLog_ pCustomLogger, 1209, LogLevel_Warning, "loadBJODroidInit", "%s: Droid %u has id = 0 - ignoring id value", path, i
        End If
        droid_.name_ = CStringOf(nameBuff, nameLength)
        droid_.position_x = (x And Not TILE_MASK) + TILE_UNITS \ 2
        droid_.position_y = (y And Not TILE_MASK) + TILE_UNITS \ 2
        ' ignore z component
        droid_.direction = DEG_u16(direction)
        droid_.player = bjoConvertPlayer(player, mapMaxPlayers)
        ' check inFire
        If inFire <> 0 Then
            printLog_ pCustomLogger, 1220, LogLevel_Warning, "loadBJODroidInit", "%s: Ignoring inFire(%i) for droid %u", path, inFire, i
        End If
        ' check periodicalDamageStart
        If periodicalDamageStart <> 0 Then
            printLog_ pCustomLogger, 1225, LogLevel_Warning, "loadBJODroidInit", "%s: Ignoring periodicalDamageStart(%u) for droid %u", path, periodicalDamageStart, i
        End If
        ' check periodicalDamage
        If periodicalDamage <> 0 Then
            printLog_ pCustomLogger, 1230, LogLevel_Warning, "loadBJODroidInit", "%s: Ignoring periodicalDamage(%u) for droid %u", path, periodicalDamage, i
        End If
        ' TODO: Sanity check droid position ?
        objects.push_back droid_
        i = i + 1
    Loop
    ' Check: extra bytes at end
    If Not pStream.endOfStream() Then
        printLog_ pCustomLogger, 1238, LogLevel_Warning, "loadBJODroidInit", "%s: Unexpectedly did not reach end of stream - data may be corrupted", path
    End If
    Set resultObjects = objects
    loadBJODroidInit = True
End Function

Private Function isKnownDroidJSONKey(ByVal key As String) As Boolean
    Select Case key
    Case "template", "id", "position", "rotation", "player", "startpos"
        isKnownDroidJSONKey = True
    End Select
End Function

Public Function loadJsonDroidInit(ByVal filename As String, ByVal mapType_ As MapType, ByVal mapIO As IOProvider, ByVal pCustomLogger As LoggingProtocol, resultObjects As ObjectVector, resultFileFormatVersion As Long) As Boolean
    Dim path As String, objects As ObjectVector, mRoot As nlohmann_json, detectedFormatVersion As Long
    Dim pDroidsRoot As nlohmann_json, droidJson As nlohmann_json
    Dim i As Long, j As Long, key As String, droid_ As Droid, player As Long
    Set objects = New ObjectVector
    path = filename

    Set mRoot = loadJsonObjectFromFile(filename, mapIO, pCustomLogger)
    If mRoot Is Nothing Then
        ' Failed to load JSON - rely on loadJsonFromFile to handle output of errors
        Exit Function
    End If

    printLog_ pCustomLogger, 1257, LogLevel_Info, "loadJsonDroidInit", "Loading: %s", path

    If Not jsonGetFileFormatVersion(filename, mRoot, pCustomLogger, 2, detectedFormatVersion) Then
        ' jsonGetFileFormatVersion should have already logged the reason
        Exit Function
    End If
    resultFileFormatVersion = detectedFormatVersion

    Set pDroidsRoot = jsonGetRootMapObjectsContainer(filename, mRoot, detectedFormatVersion, "droids")
    If pDroidsRoot Is Nothing Then
        ' jsonGetRootMapObjectsContainer should have already logged the reason
        Exit Function
    End If

    For i = 0 To pDroidsRoot.size() - 1
        key = ItemKey(pDroidsRoot, i)
        Set droidJson = pDroidsRoot.value_at(i)
        If Not droidJson.is_object() Then
            GoTo continue_
        End If
        Set droid_ = New Droid
        If Not jsonGetBaseMapObjectInfo(droid_, detectedFormatVersion, droidJson, path, key, pCustomLogger, "template") Then
            ' jsonGetBaseMapObjectInfo should have handled logging an error
            GoTo continue_
        End If
        ' the player is extracted from either "player" or "startpos" - see jsonGetPlayerFromObj
        If jsonGetPlayerFromObj(droidJson, mapType_, path, key, pCustomLogger, False, player) Then
            droid_.player = player
        Else
            ' Missing required "player" or "startpos" key for droid
            printLog_ pCustomLogger, 1298, LogLevel_Error, "loadJsonDroidInit", "%s: Missing required player/startpos key for droid: %s", path, key
            GoTo continue_
        End If

        ' Sanity check for unknown / unprocessed keys
        For j = 0 To droidJson.size() - 1
            If Not isKnownDroidJSONKey(droidJson.key_at(j)) Then
                printLog_ pCustomLogger, 1307, LogLevel_Warning, "loadJsonDroidInit", "%s: Unexpected structure key ""%s"" for structure: %s", path, droidJson.key_at(j), key
            End If
        Next

        objects.push_back droid_
continue_:
    Next

    Set resultObjects = objects
    loadJsonDroidInit = True
End Function

'--------------------------------------------------------------------------
' MARK: - Droid writing functions

Public Function writeMapDroidInitBJO(ByVal droids As ObjectVector, ByVal mapMaxPlayers As Long, ByVal filename As String, ByVal mapIO As IOProvider, ByVal bjoFileVersion As Long, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim path As String, pStream As BinaryIOStream, aFileType(0 To 3) As Byte, n As Long, ok As Boolean, quantity As Long
    Dim nameFieldLength As Long, bigZeroBuff(0 To 255) As Byte, nextAvailableDroidId As Long, i As Long
    Dim droid_ As Droid, nameBytes() As Byte, nameSize As Long, nameStrWriteLen As Long, nameOk As Boolean
    Dim droidId As Long, bjoPlayer As Long
    path = filename

    If ULT(39, bjoFileVersion) Then
        ' Currently, we only support up to version 39
        printLog_ pCustomLogger, 1326, LogLevel_Error, "writeMapDroidInitBJO", "Unsupported binary droidinit file version: %u; cannot write to: %s", bjoFileVersion, path
        Exit Function
    End If

    Set pStream = mapIO.openBinaryStream(filename, OpenMode_WRITE)

    If pStream Is Nothing Then
        printLog_ pCustomLogger, 1334, LogLevel_Error, "writeMapDroidInitBJO", "Failed to open file for writing: %s", path
        Exit Function
    End If

    printLog_ pCustomLogger, 1338, LogLevel_Info, "writeMapDroidInitBJO", "Writing: %s", path

    ' write header
    aFileType(0) = 100  ' d
    aFileType(1) = 105  ' i
    aFileType(2) = 110  ' n
    aFileType(3) = 116  ' t
    quantity = droids.size()
    ok = pStream.writeBytes(aFileType, 0, 4, n)
    If ok Then ok = (n = 4)
    If ok Then ok = BinaryIOStream_writeULE32(pStream, bjoFileVersion)
    If ok Then ok = BinaryIOStream_writeULE32(pStream, quantity)
    If Not ok Then
        printLog_ pCustomLogger, 1351, LogLevel_Error, "writeMapDroidInitBJO", "Failed writing header to: %s", path
        Exit Function
    End If

    nameFieldLength = 60
    If Not ULT(19, bjoFileVersion) Then
        nameFieldLength = 40
    End If

    nextAvailableDroidId = U32Inc(getLargestSpecifiedId(droids))

    For i = 0 To quantity - 1
        Set droid_ = droids.at(i)
        nameSize = utf8_encode(droid_.name_, nameBytes)
        nameStrWriteLen = nameFieldLength
        If nameSize < nameStrWriteLen Then nameStrWriteLen = nameSize
        If nameStrWriteLen < nameSize Then
            ' Warn about truncation!!
            printLog_ pCustomLogger, 1371, LogLevel_Warning, "writeMapDroidInitBJO", "Droids's name exceeds the length supported by binary struct file version: %u; output will be truncated", bjoFileVersion
        End If
        If Not writeBJOName(pStream, nameBytes, nameStrWriteLen, nameFieldLength, bigZeroBuff, nameOk) Then
            If Not nameOk Then
                ' Failed to write
                printLog_ pCustomLogger, 1376, LogLevel_Error, "writeMapDroidInitBJO", "Error writing structure %u's name to: %s", i, path
            Else
                ' Failed to pad with zeros
                printLog_ pCustomLogger, 1383, LogLevel_Error, "writeMapDroidInitBJO", "Error writing structure %u to: %s", i, path
            End If
            Exit Function
        End If

        droidId = 0
        If droid_.hasId Then
            droidId = droid_.id
        Else
            droidId = nextAvailableDroidId
            nextAvailableDroidId = U32Inc(nextAvailableDroidId)
        End If

        bjoPlayer = convertPlayerToBJOPlayer(droid_.player, mapMaxPlayers)
        ' TODO: Do campaign maps require scanvengers at player ID 7? (Theoretically only an issue for old campaign maps with > 7 players? And there shouldn't be any of those?)

        ok = BinaryIOStream_writeULE32(pStream, droidId)
        If ok Then ok = BinaryIOStream_writeULE32(pStream, droid_.position_x)
        If ok Then ok = BinaryIOStream_writeULE32(pStream, droid_.position_y)
        If ok Then ok = BinaryIOStream_writeULE32(pStream, 0)
        If ok Then ok = BinaryIOStream_writeULE32(pStream, droid_.direction)   ' TODO: Conversion / manipulation needed?
        If ok Then ok = BinaryIOStream_writeULE32(pStream, bjoPlayer)
        If ok Then ok = BinaryIOStream_writeSLE32(pStream, 0)   ' BOOL inFire
        If ok Then ok = BinaryIOStream_writeULE32(pStream, 0)   ' burnStart
        If ok Then ok = BinaryIOStream_writeULE32(pStream, 0)   ' burnDamage
        If Not ok Then
            printLog_ pCustomLogger, 1408, LogLevel_Error, "writeMapDroidInitBJO", "Error writing droid %u to: %s", i, path
            Exit Function
        End If
    Next

    Set pStream = Nothing
    writeMapDroidInitBJO = True
End Function

Public Function writeMapDroidInitJSON(ByVal droids As ObjectVector, ByVal mapType_ As MapType, ByVal filename As String, ByVal mapIO As IOProvider, ByVal jsonFileVersion As Long, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim path As String, mRoot As nlohmann_json, pDroidsRoot As nlohmann_json, droidObj As nlohmann_json
    Dim counter As Long, structureCountMinLength As Long, i As Long, droid_ As Droid, counterStr As String
    path = filename
    If ULT(2, jsonFileVersion) Then
        ' Currently, we only support up to version 2 JSON files
        printLog_ pCustomLogger, 1423, LogLevel_Error, "writeMapDroidInitJSON", "Unsupported JSON droidinit file version: %u; cannot write to: %s", jsonFileVersion, path
        Exit Function
    End If

    Set mRoot = json_object(True)
    Set pDroidsRoot = mRoot
    If ULT(1, jsonFileVersion) Then
        ' Version 2+ JSON format contains a "version" key
        ' And stores the objects in an array
        mRoot.set_item_ref "version", json_value_unsigned(jsonFileVersion, True)
        mRoot.set_item_ref "droids", json_array(True)
        Set pDroidsRoot = mRoot.at("droids")
    End If

    counter = 0
    structureCountMinLength = numberOfDigitsInNumber(droids.size())
    If structureCountMinLength < 1 Then structureCountMinLength = 1
    For i = 0 To droids.size() - 1
        Set droid_ = droids.at(i)
        Set droidObj = json_object(True)
        ' write base object info (name, <id>, position, rotation)
        jsonSetBaseMapObjectInfo droidObj, jsonFileVersion, droid_, "template"
        ' player
        If Not jsonSetPlayerOnObject(droidObj, mapType_, droid_.player) Then
            printLog_ pCustomLogger, 1448, LogLevel_Error, "writeMapDroidInitJSON", "Invalid player number (%i)", droid_.player
            Exit Function
        End If

        If pDroidsRoot.is_object() Then
            ' add to the root object
            counterStr = CStr(counter)
            leftPadStrToMinimumLength counterStr, structureCountMinLength, "0"
            pDroidsRoot.set_item_ref "droid_" & counterStr, droidObj
        Else 'if (pDroidsRoot->is_array())
            pDroidsRoot.push_back_ref droidObj
        End If
        counter = counter + 1
    Next

    ' write out to file
    writeMapDroidInitJSON = saveOrderedJsonObjectToFile(mRoot, filename, mapIO)
End Function

Public Function writeMapDroidInit(ByVal droids As ObjectVector, ByVal mapMaxPlayers As Long, ByVal mapType_ As MapType, ByVal mapFolderPath As String, ByVal mapIO As IOProvider, ByVal format As OutputFormat, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim retVal As Boolean
    retVal = False
    Select Case format
    Case OutputFormat_VER1_BINARY_OLD
        retVal = writeMapDroidInitBJO(droids, mapMaxPlayers, IOProvider_pathJoin(mapIO, mapFolderPath, "dinit.bjo"), mapIO, 8, pCustomLogger)
    Case OutputFormat_VER2
        retVal = writeMapDroidInitJSON(droids, mapType_, IOProvider_pathJoin(mapIO, mapFolderPath, "droid.json"), mapIO, 1, pCustomLogger)
    Case OutputFormat_VER3
        retVal = writeMapDroidInitJSON(droids, mapType_, IOProvider_pathJoin(mapIO, mapFolderPath, "droid.json"), mapIO, 2, pCustomLogger)
    End Select
    writeMapDroidInit = retVal
End Function

'--------------------------------------------------------------------------
' MARK: - Feature loading functions

Public Function loadBJOFeatureInit(ByVal filename As String, ByVal mapMaxPlayers As Long, ByVal mapIO As IOProvider, ByVal pCustomLogger As LoggingProtocol, resultObjects As ObjectVector, resultFileFormatVersion As Long) As Boolean
    Dim path As String, pStream As BinaryIOStream, objects As ObjectVector
    Dim aFileType(0 To 3) As Byte, n As Long, ok As Boolean, version As Long, quantity As Long
    Dim nameLength As Long, nameBuff() As Byte, i As Long, visIdx As Long, visibility(0 To 7) As Byte
    Dim id As Long, x As Long, y As Long, z As Long, direction As Long, player As Long
    Dim inFire As Long, periodicalDamageStart As Long, periodicalDamage As Long, feature_ As Feature, converted_player As Long
    Set objects = New ObjectVector
    path = filename

    Set pStream = mapIO.openBinaryStream(filename, OpenMode_READ)
    If pStream Is Nothing Then
        ' Failed to open file
        Exit Function
    End If

    printLog_ pCustomLogger, 1502, LogLevel_Info, "loadBJOFeatureInit", "Loading: %s", path

    version = 0
    quantity = 0
    ok = pStream.readBytes(aFileType, 0, 4, n)
    If ok Then ok = (n = 4)
    If ok Then ok = (aFileType(0) = 102 And aFileType(1) = 101 And aFileType(2) = 97 And aFileType(3) = 116)   ' "feat"
    If ok Then ok = BinaryIOStream_readULE32(pStream, version)
    If ok Then ok = BinaryIOStream_readULE32(pStream, quantity)
    If Not ok Then
        printLog_ pCustomLogger, 1515, LogLevel_Error, "loadBJOFeatureInit", "%s: Bad header", path
        Exit Function
    End If
    resultFileFormatVersion = version

    If ULT(version, 7) Or ULT(19, version) Then
        printLog_ pCustomLogger, 1522, LogLevel_Error, "loadBJOFeatureInit", "%s: Unsupported binary feat file version: %u", path, version
        Exit Function
    End If

    nameLength = 60
    If Not ULT(19, version) Then
        nameLength = 40
    End If
    ReDim nameBuff(0 To nameLength - 1)

    i = 0
    Do While ULT(i, quantity)
        id = 0
        x = 0: y = 0: z = 0
        direction = 0
        player = 0

        inFire = 0
        periodicalDamageStart = 0
        periodicalDamage = 0

        ok = pStream.readBytes(nameBuff, 0, nameLength, n)
        If ok Then ok = (n = nameLength)
        If ok Then ok = BinaryIOStream_readULE32(pStream, id)
        If ok Then ok = BinaryIOStream_readULE32(pStream, x)
        If ok Then ok = BinaryIOStream_readULE32(pStream, y)
        If ok Then ok = BinaryIOStream_readULE32(pStream, z)
        If ok Then ok = BinaryIOStream_readULE32(pStream, direction)
        If ok Then ok = BinaryIOStream_readULE32(pStream, player)
        If ok Then ok = BinaryIOStream_readSLE32(pStream, inFire)                  ' BOOL inFire
        If ok Then ok = BinaryIOStream_readULE32(pStream, periodicalDamageStart)   ' burnStart
        If ok Then ok = BinaryIOStream_readULE32(pStream, periodicalDamage)        ' burnDamage
        If Not ok Then
            printLog_ pCustomLogger, 1554, LogLevel_Error, "loadBJOFeatureInit", "%s: Failed to read feature %u", path, i
            Exit Function
        End If
        If Not ULT(version, 14) Then
            ok = pStream.readBytes(visibility, 0, 8, n)
            If ok Then ok = (n = 8)
            If Not ok Then
                printLog_ pCustomLogger, 1559, LogLevel_Error, "loadBJOFeatureInit", "%s: Failed to read feature %u visibility", path, i
                Exit Function
            End If
        End If

        Set feature_ = New Feature
        If id <> 0 Then
            feature_.id = id
            feature_.hasId = True
        Else
            printLog_ pCustomLogger, 1570, LogLevel_Warning, "loadBJOFeatureInit", "%s: Feature %u has id = 0 - ignoring id value", path, i
        End If
        feature_.name_ = CStringOf(nameBuff, nameLength)
        feature_.position_x = x
        feature_.position_y = y
        ' ignore z component
        feature_.direction = DEG_u16(direction)
        ' check player - ONLY POSSIBLY USED FOR CAMPAIGN, but prior code always ignored it??
        converted_player = bjoConvertPlayer(player, mapMaxPlayers)
        If converted_player <> PLAYER_SCAVENGERS And converted_player <> mapMaxPlayers Then
            ' Since the prior code did not actually utilize this value, just check and print a warning (for now) if we are ignoring it
            printLog_ pCustomLogger, 1582, LogLevel_Warning, "loadBJOFeatureInit", "%s: Ignoring player(%i) for feature %u", path, player, i
        End If
        ' check inFire
        If inFire <> 0 Then
            printLog_ pCustomLogger, 1587, LogLevel_Warning, "loadBJOFeatureInit", "%s: Ignoring inFire(%i) for feature %u", path, inFire, i
        End If
        ' check periodicalDamageStart
        If periodicalDamageStart <> 0 Then
            printLog_ pCustomLogger, 1592, LogLevel_Warning, "loadBJOFeatureInit", "%s: Ignoring periodicalDamageStart(%u) for feature %u", path, periodicalDamageStart, i
        End If
        ' check periodicalDamage
        If periodicalDamage <> 0 Then
            printLog_ pCustomLogger, 1597, LogLevel_Warning, "loadBJOFeatureInit", "%s: Ignoring periodicalDamage(%u) for feature %u", path, periodicalDamage, i
        End If
        If Not ULT(version, 14) Then
            For visIdx = 0 To 7
                If visibility(visIdx) <> 0 Then
                    printLog_ pCustomLogger, 1605, LogLevel_Warning, "loadBJOFeatureInit", "%s: Ignoring non-0 visibility[%zu]=%u for feature %u", path, visIdx, visibility(visIdx), i
                End If
            Next
        End If
        ' TODO: Sanity check feature position ?
        objects.push_back feature_
        i = i + 1
    Loop
    ' Check: extra bytes at end
    If Not pStream.endOfStream() Then
        printLog_ pCustomLogger, 1615, LogLevel_Warning, "loadBJOFeatureInit", "%s: Unexpectedly did not reach end of stream - data may be corrupted", path
    End If
    Set resultObjects = objects
    loadBJOFeatureInit = True
End Function

Private Function isKnownFeatureJSONKey(ByVal key As String) As Boolean
    Select Case key
    Case "name", "id", "position", "rotation", "player"
        isKnownFeatureJSONKey = True
    End Select
End Function

Public Function loadJsonFeatureInit(ByVal filename As String, ByVal mapType_ As MapType, ByVal mapIO As IOProvider, ByVal pCustomLogger As LoggingProtocol, resultObjects As ObjectVector, resultFileFormatVersion As Long) As Boolean
    Dim path As String, objects As ObjectVector, mRoot As nlohmann_json, detectedFormatVersion As Long
    Dim pFeaturesRoot As nlohmann_json, featureJson As nlohmann_json
    Dim i As Long, j As Long, key As String, feature_ As Feature, player As Long
    Set objects = New ObjectVector
    path = filename

    Set mRoot = loadJsonObjectFromFile(filename, mapIO, pCustomLogger)
    If mRoot Is Nothing Then
        ' Failed to load JSON - rely on loadJsonFromFile to handle output of errors
        Exit Function
    End If

    printLog_ pCustomLogger, 1634, LogLevel_Info, "loadJsonFeatureInit", "Loading: %s", path

    If Not jsonGetFileFormatVersion(filename, mRoot, pCustomLogger, 2, detectedFormatVersion) Then
        ' jsonGetFileFormatVersion should have already logged the reason
        Exit Function
    End If
    resultFileFormatVersion = detectedFormatVersion

    Set pFeaturesRoot = jsonGetRootMapObjectsContainer(filename, mRoot, detectedFormatVersion, "features")
    If pFeaturesRoot Is Nothing Then
        ' jsonGetRootMapObjectsContainer should have already logged the reason
        Exit Function
    End If

    For i = 0 To pFeaturesRoot.size() - 1
        key = ItemKey(pFeaturesRoot, i)
        Set featureJson = pFeaturesRoot.value_at(i)
        If Not featureJson.is_object() Then
            GoTo continue_
        End If
        Set feature_ = New Feature
        If Not jsonGetBaseMapObjectInfo(feature_, detectedFormatVersion, featureJson, path, key, pCustomLogger, "name") Then
            ' jsonGetBaseMapObjectInfo should have handled logging an error
            GoTo continue_
        End If
        ' Optional:
        ' the player is extracted from "player" *only*
        ' FIXME: Is this actually used for skirmish map feature init?
        If jsonGetPlayerFromObj(featureJson, mapType_, path, key, pCustomLogger, True, player) Then
            If mapType_ = MapType_CAMPAIGN Or mapType_ = MapType_SAVEGAME Then
                feature_.player = player
                feature_.hasPlayer = True
            Else
                ' player assignment for features is not expected for skirmish map init
                printLog_ pCustomLogger, 1679, LogLevel_Warning, "loadJsonFeatureInit", "%s: Ignoring assigned player (%u) for feature: %s", path, player, key
            End If
        End If

        ' Sanity check for unknown / unprocessed keys
        For j = 0 To featureJson.size() - 1
            If Not isKnownFeatureJSONKey(featureJson.key_at(j)) Then
                printLog_ pCustomLogger, 1688, LogLevel_Warning, "loadJsonFeatureInit", "%s: Unexpected structure key ""%s"" for structure: %s", path, featureJson.key_at(j), key
            End If
        Next

        objects.push_back feature_
continue_:
    Next

    Set resultObjects = objects
    loadJsonFeatureInit = True
End Function

'--------------------------------------------------------------------------
' MARK: - Feature writing functions

Public Function writeMapFeatureInitBJO(ByVal features As ObjectVector, ByVal mapMaxPlayers As Long, ByVal filename As String, ByVal mapIO As IOProvider, ByVal bjoFileVersion As Long, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim path As String, pStream As BinaryIOStream, aFileType(0 To 3) As Byte, n As Long, ok As Boolean, quantity As Long
    Dim nameFieldLength As Long, bigZeroBuff(0 To 255) As Byte, nextAvailableFeatureId As Long, i As Long
    Dim feature_ As Feature, nameBytes() As Byte, nameSize As Long, nameStrWriteLen As Long, nameOk As Boolean
    Dim featureId As Long, bjoPlayer As Long
    path = filename

    If ULT(bjoFileVersion, 7) Or ULT(19, bjoFileVersion) Then
        ' Currently, we only support versions 7 - 19
        printLog_ pCustomLogger, 1707, LogLevel_Error, "writeMapFeatureInitBJO", "Unsupported binary feat file version: %u; cannot write to: %s", bjoFileVersion, path
        Exit Function
    End If

    Set pStream = mapIO.openBinaryStream(filename, OpenMode_WRITE)

    If pStream Is Nothing Then
        printLog_ pCustomLogger, 1715, LogLevel_Error, "writeMapFeatureInitBJO", "Failed to open file for writing: %s", path
        Exit Function
    End If

    printLog_ pCustomLogger, 1719, LogLevel_Info, "writeMapFeatureInitBJO", "Writing: %s", path

    ' write header
    aFileType(0) = 102  ' f
    aFileType(1) = 101  ' e
    aFileType(2) = 97   ' a
    aFileType(3) = 116  ' t
    quantity = features.size()
    ok = pStream.writeBytes(aFileType, 0, 4, n)
    If ok Then ok = (n = 4)
    If ok Then ok = BinaryIOStream_writeULE32(pStream, bjoFileVersion)
    If ok Then ok = BinaryIOStream_writeULE32(pStream, quantity)
    If Not ok Then
        printLog_ pCustomLogger, 1732, LogLevel_Error, "writeMapFeatureInitBJO", "Failed writing header to: %s", path
        Exit Function
    End If

    nameFieldLength = 60
    If Not ULT(19, bjoFileVersion) Then
        nameFieldLength = 40
    End If

    nextAvailableFeatureId = U32Inc(getLargestSpecifiedId(features))

    For i = 0 To quantity - 1
        Set feature_ = features.at(i)
        nameSize = utf8_encode(feature_.name_, nameBytes)
        nameStrWriteLen = nameFieldLength
        If nameSize < nameStrWriteLen Then nameStrWriteLen = nameSize
        If nameStrWriteLen < nameSize Then
            ' Warn about truncation!!
            printLog_ pCustomLogger, 1752, LogLevel_Warning, "writeMapFeatureInitBJO", "Droids's name exceeds the length supported by binary struct file version: %u; output will be truncated", bjoFileVersion
        End If
        If Not writeBJOName(pStream, nameBytes, nameStrWriteLen, nameFieldLength, bigZeroBuff, nameOk) Then
            If Not nameOk Then
                ' Failed to write
                printLog_ pCustomLogger, 1757, LogLevel_Error, "writeMapFeatureInitBJO", "Error writing structure %u's name to: %s", i, path
            Else
                ' Failed to pad with zeros
                printLog_ pCustomLogger, 1764, LogLevel_Error, "writeMapFeatureInitBJO", "Error writing structure %u to: %s", i, path
            End If
            Exit Function
        End If

        featureId = 0
        If feature_.hasId Then
            featureId = feature_.id
        Else
            featureId = nextAvailableFeatureId
            nextAvailableFeatureId = U32Inc(nextAvailableFeatureId)
        End If

        bjoPlayer = mapMaxPlayers
        If feature_.hasPlayer Then
            bjoPlayer = convertPlayerToBJOPlayer(feature_.player, mapMaxPlayers)
        End If
        ' TODO: Do campaign maps require scanvengers at player ID 7? (Theoretically only an issue for old campaign maps with > 7 players? And there shouldn't be any of those?)

        ok = BinaryIOStream_writeULE32(pStream, featureId)
        If ok Then ok = BinaryIOStream_writeULE32(pStream, feature_.position_x)
        If ok Then ok = BinaryIOStream_writeULE32(pStream, feature_.position_y)
        If ok Then ok = BinaryIOStream_writeULE32(pStream, 0)
        If ok Then ok = BinaryIOStream_writeULE32(pStream, feature_.direction)    ' TODO: Conversion / manipulation needed?
        If ok Then ok = BinaryIOStream_writeULE32(pStream, bjoPlayer)
        If ok Then ok = BinaryIOStream_writeSLE32(pStream, 0)   ' BOOL inFire
        If ok Then ok = BinaryIOStream_writeULE32(pStream, 0)   ' burnStart
        If ok Then ok = BinaryIOStream_writeULE32(pStream, 0)   ' burnDamage
        If Not ok Then
            printLog_ pCustomLogger, 1793, LogLevel_Error, "writeMapFeatureInitBJO", "Error writing feature %u to: %s", i, path
            Exit Function
        End If
        If Not ULT(bjoFileVersion, 14) Then
            ok = pStream.writeBytes(bigZeroBuff, 0, 8, n)
            If ok Then ok = (n = 8)
            If Not ok Then
                printLog_ pCustomLogger, 1798, LogLevel_Error, "writeMapFeatureInitBJO", "Error writing feature %u visibility to: %s", i, path
                Exit Function
            End If
        End If
    Next

    Set pStream = Nothing
    writeMapFeatureInitBJO = True
End Function

Public Function writeMapFeatureInitJSON(ByVal features As ObjectVector, ByVal mapType_ As MapType, ByVal filename As String, ByVal mapIO As IOProvider, ByVal jsonFileVersion As Long, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim path As String, mRoot As nlohmann_json, pFeaturesRoot As nlohmann_json, featureObj As nlohmann_json
    Dim counter As Long, structureCountMinLength As Long, i As Long, feature_ As Feature, counterStr As String
    path = filename
    If ULT(2, jsonFileVersion) Then
        ' Currently, we only support up to version 2 JSON files
        printLog_ pCustomLogger, 1813, LogLevel_Error, "writeMapFeatureInitJSON", "Unsupported JSON feature file version: %u; cannot write to: %s", jsonFileVersion, path
        Exit Function
    End If

    Set mRoot = json_object(True)
    Set pFeaturesRoot = mRoot
    If ULT(1, jsonFileVersion) Then
        ' Version 2+ JSON format contains a "version" key
        ' And stores the objects in an array
        mRoot.set_item_ref "version", json_value_unsigned(jsonFileVersion, True)
        mRoot.set_item_ref "features", json_array(True)
        Set pFeaturesRoot = mRoot.at("features")
    End If

    counter = 0
    structureCountMinLength = numberOfDigitsInNumber(features.size())
    If structureCountMinLength < 1 Then structureCountMinLength = 1
    For i = 0 To features.size() - 1
        Set feature_ = features.at(i)
        Set featureObj = json_object(True)
        ' write base object info (name, <id>, position, rotation)
        jsonSetBaseMapObjectInfo featureObj, jsonFileVersion, feature_, "name"
        ' Optional:
        ' the player is extracted from "player" *only*
        If feature_.hasPlayer Then
            If mapType_ = MapType_CAMPAIGN Or mapType_ = MapType_SAVEGAME Then
                If Not jsonSetPlayerOnObject(featureObj, mapType_, feature_.player) Then
                    printLog_ pCustomLogger, 1843, LogLevel_Error, "writeMapFeatureInitJSON", "Invalid player number (%i)", feature_.player
                    Exit Function
                End If
            Else
                ' player assignment for features is not expected for skirmish map init
                printLog_ pCustomLogger, 1850, LogLevel_Warning, "writeMapFeatureInitJSON", "Ignoring assigned player (%u) for feature #%zu", feature_.player, counter
            End If
        End If

        If pFeaturesRoot.is_object() Then
            ' add to the root object
            counterStr = CStr(counter)
            leftPadStrToMinimumLength counterStr, structureCountMinLength, "0"
            pFeaturesRoot.set_item_ref "feature_" & counterStr, featureObj
        Else 'if (pFeaturesRoot->is_array())
            pFeaturesRoot.push_back_ref featureObj
        End If
        counter = counter + 1
    Next

    ' write out to file
    writeMapFeatureInitJSON = saveOrderedJsonObjectToFile(mRoot, filename, mapIO)
End Function

Public Function writeMapFeatureInit(ByVal features As ObjectVector, ByVal mapMaxPlayers As Long, ByVal mapType_ As MapType, ByVal mapFolderPath As String, ByVal mapIO As IOProvider, ByVal format As OutputFormat, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim retVal As Boolean
    retVal = False
    Select Case format
    Case OutputFormat_VER1_BINARY_OLD
        retVal = writeMapFeatureInitBJO(features, mapMaxPlayers, IOProvider_pathJoin(mapIO, mapFolderPath, "feat.bjo"), mapIO, 8, pCustomLogger)
    Case OutputFormat_VER2
        retVal = writeMapFeatureInitJSON(features, mapType_, IOProvider_pathJoin(mapIO, mapFolderPath, "feature.json"), mapIO, 1, pCustomLogger)
    Case OutputFormat_VER3
        retVal = writeMapFeatureInitJSON(features, mapType_, IOProvider_pathJoin(mapIO, mapFolderPath, "feature.json"), mapIO, 2, pCustomLogger)
    End Select
    writeMapFeatureInit = retVal
End Function

'--------------------------------------------------------------------------
' MARK: - High-level Map implementation (the static members of class Map)

' The mapIO argument: omitted is a new StdIOProvider (the C++ default), Nothing is nullptr
Private Function MapIOArgument(mapIO As Variant) As IOProvider
    If IsMissing(mapIO) Then
        Set MapIOArgument = New StdIOProvider
    ElseIf IsObject(mapIO) Then
        Set MapIOArgument = mapIO
    End If
End Function

' Construct an empty Map, for modification (Map::Map())
Public Function Map_new() As Map
    Dim m As Map
    Set m = New Map
    m.construct_empty
    Set Map_new = m
End Function

' Load a map from a specified folder path + mapType + maxPlayers + random seed (only used for script-generated maps), optionally supplying:
' - a logger
' - a WzMap::IOProvider
Public Function Map_loadFromPath(ByVal mapFolderPath As String, ByVal mapType_ As MapType, ByVal mapMaxPlayers As Long, ByVal seed As Long, Optional ByVal logger As LoggingProtocol = Nothing, Optional ByVal mapIO As Variant) As Map
    Dim io As IOProvider, fileData() As Byte, fileSize As Long, gameJSPath As String, gameJSLoadResult As LoadFullFileResult
    Dim result As Map, pMap As Map
    Set io = MapIOArgument(mapIO)
    If io Is Nothing Then
        printLog_ logger, 1919, LogLevel_Error, "loadFromPath", "Missing required mapIO"
        Exit Function
    End If

    ' First, check for new game.js format
    gameJSPath = IOProvider_pathJoin(io, mapFolderPath, "game.js")
    gameJSLoadResult = io.loadFullFile(gameJSPath, fileData, fileSize, MaxGameJSFileSize, True)
    Select Case gameJSLoadResult
    Case LoadFullFileResult_SUCCESS
        If ULT(MaxGameJSFileSize, fileSize) Then ' double-check (in case mapIO doesn't handle maxFileSize)
            printLog_ logger, 1933, LogLevel_Error, "loadFromPath", "game.js file is too large: %s", gameJSPath
            Exit Function
        End If

        printLog_ logger, 1937, LogLevel_Info, "loadFromPath", "Loading: %s", gameJSPath
        ' Load script map, which actually loads everything
        Set result = runMapScript(fileData, fileSize, gameJSPath, seed, False, logger)
        If Not result Is Nothing Then
            result.setScriptGenerated mapFolderPath, mapType_, mapMaxPlayers, io, logger, seed, fileData, fileSize, Nothing
        End If
        Set Map_loadFromPath = result
        Exit Function
    Case LoadFullFileResult_FAILURE_OPEN
        ' unable to open game.js file - fallback to lazy-loading regular map
    Case LoadFullFileResult_FAILURE_READ
        printLog_ logger, 1958, LogLevel_Error, "loadFromPath", "Failed to read game.js file: %s", gameJSPath
        Exit Function
    Case LoadFullFileResult_FAILURE_EXCEEDS_MAXFILESIZE
        printLog_ logger, 1961, LogLevel_Error, "loadFromPath", "game.js file is too large: %s", gameJSPath
        Exit Function
    End Select

    ' Otherwise, construct a lazy-loading Map
    Set pMap = New Map
    pMap.construct mapFolderPath, mapType_, mapMaxPlayers, logger, io
    Set Map_loadFromPath = pMap
End Function

' Export a map to a specified folder path in a specified output format (version)
Public Function Map_exportMapToPath(ByVal map_ As Map, ByVal mapFolderPath As String, ByVal mapType_ As MapType, ByVal mapMaxPlayers As Long, ByVal format As OutputFormat, Optional ByVal logger As LoggingProtocol = Nothing, Optional ByVal mapIO As Variant) As Boolean
    Dim io As IOProvider, pMapData As MapData, pMapTerrainTypes As TerrainTypeData
    Dim pMapDroids As ObjectVector, pMapFeatures As ObjectVector, pMapStructures As ObjectVector
    Set io = MapIOArgument(mapIO)
    If io Is Nothing Then
        printLog_ logger, 1974, LogLevel_Error, "exportMapToPath", "Missing required mapIO"
        Exit Function
    End If

    Set pMapData = map_.mapData()
    If pMapData Is Nothing Then
        printLog_ logger, 1981, LogLevel_Error, "exportMapToPath", "Failed to load / retrieve map data from: %s", map_.mapFolderPath()
        Exit Function
    End If
    If Not writeMapData(pMapData, IOProvider_pathJoin(io, mapFolderPath, "game.map"), io, format, logger) Then
        printLog_ logger, 1986, LogLevel_Error, "exportMapToPath", "Failed to write map data to path: %s", mapFolderPath
        Exit Function
    End If

    Set pMapTerrainTypes = map_.mapTerrainTypes()
    If pMapTerrainTypes Is Nothing Then
        printLog_ logger, 1993, LogLevel_Error, "exportMapToPath", "Failed to load / retrieve map terrain type data from: %s", map_.mapFolderPath()
        Exit Function
    End If
    If Not writeTerrainTypes(pMapTerrainTypes, IOProvider_pathJoin(io, mapFolderPath, "ttypes.ttp"), io, format, logger) Then
        printLog_ logger, 1998, LogLevel_Error, "exportMapToPath", "Failed to write map terrain type data to path: %s", mapFolderPath
        Exit Function
    End If

    Set pMapDroids = map_.mapDroids()
    If pMapDroids Is Nothing Then
        printLog_ logger, 2005, LogLevel_Error, "exportMapToPath", "Failed to load / retrieve map droids from: %s", map_.mapFolderPath()
        Exit Function
    End If
    If Not writeMapDroidInit(pMapDroids, mapMaxPlayers, mapType_, mapFolderPath, io, format, logger) Then
        printLog_ logger, 2010, LogLevel_Error, "exportMapToPath", "Failed to write map droids to path: %s", mapFolderPath
        Exit Function
    End If

    Set pMapFeatures = map_.mapFeatures()
    If pMapFeatures Is Nothing Then
        printLog_ logger, 2017, LogLevel_Error, "exportMapToPath", "Failed to load / retrieve map features from: %s", map_.mapFolderPath()
        Exit Function
    End If
    If Not writeMapFeatureInit(pMapFeatures, mapMaxPlayers, mapType_, mapFolderPath, io, format, logger) Then
        printLog_ logger, 2022, LogLevel_Error, "exportMapToPath", "Failed to write map features to path: %s", mapFolderPath
        Exit Function
    End If

    Set pMapStructures = map_.mapStructures()
    If pMapStructures Is Nothing Then
        printLog_ logger, 2029, LogLevel_Error, "exportMapToPath", "Failed to load / retrieve map structures from: %s", map_.mapFolderPath()
        Exit Function
    End If
    If Not writeMapStructureInit(pMapStructures, mapMaxPlayers, mapType_, mapFolderPath, io, format, logger) Then
        printLog_ logger, 2034, LogLevel_Error, "exportMapToPath", "Failed to write map structures to path: %s", mapFolderPath
        Exit Function
    End If

    Map_exportMapToPath = True
End Function

Private Sub InsertName(results As StringVector, ByVal s As String)
    If Not StringVector_contains(results, s) Then StringVector_push_back results, s
End Sub

' Getting a list of expected file names for a particular map format (an
' unordered_set: each name once, in names(0 .. count - 1))
Public Function Map_expectedFileNames(names() As String, ByVal hasFormat As Boolean, ByVal format As LoadedFormat) As Long
    Dim results As StringVector, f As LoadedFormat, i As Long
    If hasFormat And format = LoadedFormat_MIXED Then
        hasFormat = False
    End If
    If Not hasFormat Then
        ' ancient .ini files
        InsertName results, "droid.ini": InsertName results, "feature.ini": InsertName results, "struct.ini"
    End If
    If hasFormat Then f = format Else f = LoadedFormat_BINARY_OLD
    If f = LoadedFormat_BINARY_OLD Then
        InsertName results, "game.map": InsertName results, "ttypes.ttp": InsertName results, "dinit.bjo": InsertName results, "feat.bjo": InsertName results, "struct.bjo"
    End If
    If hasFormat Then f = format Else f = LoadedFormat_JSON_v1
    If f = LoadedFormat_JSON_v1 Then
        InsertName results, "game.map": InsertName results, "ttypes.ttp": InsertName results, "droid.json": InsertName results, "feature.json": InsertName results, "struct.json"
    End If
    If hasFormat Then f = format Else f = LoadedFormat_SCRIPT_GENERATED
    If f = LoadedFormat_SCRIPT_GENERATED Then
        InsertName results, "game.js": InsertName results, "ttypes.ttp"
    End If
    If hasFormat Then f = format Else f = LoadedFormat_JSON_v2
    If f = LoadedFormat_JSON_v2 Then
        InsertName results, "game.map": InsertName results, "ttypes.ttp": InsertName results, "droid.json": InsertName results, "feature.json": InsertName results, "struct.json"
    End If

    If results.size > 0 Then
        ReDim names(0 To results.size - 1)
        For i = 0 To results.size - 1
            names(i) = results.items(i)
        Next
    End If
    Map_expectedFileNames = results.size
End Function
