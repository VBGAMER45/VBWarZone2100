Attribute VB_Name = "map_c"
' src/map.cpp, map.h (ported in part): the map of the game: the tiles
' (psMapTiles), the tileset's ground types and decals, the riverbed, the aux
' and blocking maps, the continents, and the height and line queries. A
' MAPTILE * is a tile index (map_h). The tileset files are read from
' fileLoadBuffer (init_c) as C reads them: as text up to the terminating
' zero, a const char * being a 1-based position in it. WzMapPhysFSIO,
' WzMapBinaryPhysFSStream and WzMapDebugLogger are classes of their own.
' The danger maps (dangerFloodFill, threatUpdate, mapInit, mapUpdate) have
' no thread: the danger thread's work (dangerFloodFill of lastDangerPlayer,
' on the shadow copies auxMapStore made) is done where C++ starts the
' thread or posts dangerSemaphore, and waiting for dangerDoneSemaphore is
' nothing; the result is the same, as the thread only reads and writes the
' copies. Not ported yet: mapObjIsAboveGround (cluster B); readVisibilityData, writeVisibilityData and
' mapSaveToWzMapData (save games, cluster I); syncLogDumpAuxMaps (#if 0).
Option Explicit

' Maximum expected return value from get height
Private Const MAX_HEIGHT As Long = 256 * ELEVATION_SCALE

Private Const WATER_MIN_DEPTH As Long = 500
Private Const WATER_MAX_DEPTH As Long = WATER_MIN_DEPTH + 400

Private Const MAX_TERRAIN_TILES As Long = &H200     ' max that we support (for now), see TILE_NUMMASK

' (VB) the sscanf formats of the tileset files ("\r\n" in C's)
Private Const FMT_NAME_COUNT As String = "%255[^,'" & vbCr & vbLf & "],%d%n"
Private Const FMT_NAME As String = "%255[^,'" & vbCr & vbLf & "]%n"
Private Const FMT_TERTILE As String = "%255[^,'" & vbCr & vbLf & "],%255[^,'" & vbCr & vbLf & "],%lf%n"
Private Const FMT_GROUND As String = "%255[^,'" & vbCr & vbLf & "],%255[^,'" & vbCr & vbLf & "],%255[^,'" & vbCr & vbLf & "],%255[^,'" & vbCr & vbLf & "]%n"

'scroll min and max values
Public scrollMinX As Long, scrollMaxX As Long, scrollMinY As Long, scrollMaxY As Long

'For saves to determine if loading the terrain type override should occur
Public builtInMap As Boolean
Public useTerrainOverrides As Boolean

' The size and contents of the map
Public mapWidth As Long, mapHeight As Long
Public psMapTiles() As MAPTILE_
Public psMapTiles_allocated As Boolean  ' (VB) psMapTiles != nullptr
Public psBlockMap() As Byte             ' (slot, tile): psBlockMap[AUX_MAX]
Public psAuxMap() As Byte               ' (player, tile): psAuxMap[MAX_PLAYERS + AUX_MAX]; yes, we waste one element... eyes wide open... makes API nicer
Private auxMaps_allocated As Boolean

' The different ground types
Private groundTypes() As GROUND_TYPE
Private groundTypes_size As Long
Public tilesetDir As String             ' char *: vbNullString = nullptr
Public currentMapTileset As MAP_TILESET
Private numTile_names As Long
Private Tile_names() As String

Private map_() As Long                  ' 3D array pointer that holds the texturetype
Private numTile_types As Long

Private mapDecals() As Boolean          ' array that tells us what tile is a decal
Private mapDecals_allocated As Boolean

' Look up table that returns the terrain type of a given tile texture
Public terrainTypes(0 To MAX_TILE_TEXTURES - 1) As Byte

' Convert a direction into an offset.
' dir 0 => x = 0, y = -1
Private Const NUM_DIR As Long = 8
Private aDirOffset_x As Variant, aDirOffset_y As Variant

' the danger maps
Private Const GAME_TICKS_FOR_DANGER As Long = GAME_TICKS_PER_SEC * 2
Private floodbucket_x() As Long          ' struct floodtile *floodbucket: uint8_t x, y
Private floodbucket_y() As Long
Private bucketcounter As Long
Private lastDangerUpdate As Long        ' UDWORD
Private lastDangerPlayer As Long

Public Sub map_c_init()
    currentMapTileset = MAP_TILESET_ARIZONA
    tilesetDir = vbNullString
    aDirOffset_x = Array(0, -1, -1, -1, 0, 1, 1, 1)
    aDirOffset_y = Array(1, 1, 0, -1, -1, -1, 0, 1)
    lastDangerPlayer = -1
End Sub

'--------------------------------------------------------------------------
' (VB) reading the files in fileLoadBuffer

' the file in fileLoadBuffer as C reads it: its characters up to the terminating zero
Private Function fileLoadBufferText(ByVal fileSize As Long) As String
    Dim i As Long, s As String
    s = String$(fileSize, vbNullChar)
    For i = 0 To fileSize - 1
        If fileLoadBuffer(i) = 0 Then
            fileLoadBufferText = Left$(s, i)
            Exit Function
        End If
        Mid$(s, i + 1, 1) = ChrW$(fileLoadBuffer(i))
    Next
    fileLoadBufferText = s
End Function

' strchr(p, '\n') + 1 (where there is no newline C reads through a null pointer: here the end of the text)
Private Function nextLine(text As String, ByVal p As Long) As Long
    Dim k As Long
    k = InStr(p, text, vbLf)
    If k = 0 Then nextLine = Len(text) + 1 Else nextLine = k + 1
End Function

'--------------------------------------------------------------------------

Public Function getGroundType(ByVal idx As Long) As GROUND_TYPE
    getGroundType = groundTypes(idx)
End Function

Public Function getNumGroundTypes() As Long
    getNumGroundTypes = groundTypes_size
End Function

Private Sub init_tileNames(ByVal typ As MAP_TILESET)
    Dim text As String, p As Long, name_ As String, numlines As Long, i As Long, cnt As Long, fileSize As Long

    Select Case typ
    Case MAP_TILESET_ARIZONA
        If Not loadFileToBuffer("tileset/arizona_enum.txt", fileLoadBuffer, FILE_LOAD_BUFFER_SIZE, fileSize) Then
            debug_ 203, LOG_FATAL, "init_tileNames", "tileset/arizona_enum.txt not found.  Aborting."
            exit_ 3                     ' abort()
        End If
        text = fileLoadBufferText(fileSize)
        p = 1

        sscanf Mid$(text, p), FMT_NAME_COUNT, name_, numlines, cnt
        p = p + cnt

        If name_ <> "arizona_enum" Then
            debug_ 212, LOG_FATAL, "init_tileNames", "%s found, but was expecting arizona_enum, aborting.", name_
            exit_ 3                     ' abort()
        End If
    Case MAP_TILESET_URBAN
        If Not loadFileToBuffer("tileset/urban_enum.txt", fileLoadBuffer, FILE_LOAD_BUFFER_SIZE, fileSize) Then
            debug_ 221, LOG_FATAL, "init_tileNames", "tileset/urban_enum.txt not found.  Aborting."
            exit_ 3                     ' abort()
        End If
        text = fileLoadBufferText(fileSize)
        p = 1

        sscanf Mid$(text, p), FMT_NAME_COUNT, name_, numlines, cnt
        p = p + cnt

        If name_ <> "urban_enum" Then
            debug_ 230, LOG_FATAL, "init_tileNames", "%s found, but was expecting urban_enum, aborting.", name_
            exit_ 3                     ' abort()
        End If
    Case MAP_TILESET_ROCKIES
        If Not loadFileToBuffer("tileset/rockie_enum.txt", fileLoadBuffer, FILE_LOAD_BUFFER_SIZE, fileSize) Then
            debug_ 239, LOG_FATAL, "init_tileNames", "tileset/rockie_enum.txt not found.  Aborting."
            exit_ 3                     ' abort()
        End If
        text = fileLoadBufferText(fileSize)
        p = 1

        sscanf Mid$(text, p), FMT_NAME_COUNT, name_, numlines, cnt
        p = p + cnt

        If name_ <> "rockie_enum" Then
            debug_ 248, LOG_FATAL, "init_tileNames", "%s found, but was expecting rockie_enum, aborting.", name_
            exit_ 3                     ' abort()
        End If
    End Select

    debug_ 255, LOG_TERRAIN, "init_tileNames", "name: %s, with %d entries", name_, numlines
    If numlines = 0 Or numlines > MAX_TERRAIN_TILES Then
        debug_ 258, LOG_FATAL, "init_tileNames", "Rockie_enum parameter is out of range (%d). Aborting.", numlines
        exit_ 3                         ' abort()
    End If

    numTile_names = numlines
    'increment the pointer to the start of the next record
    p = nextLine(text, p)
    ReDim Tile_names(0 To numlines - 1)

    For i = 0 To numlines - 1
        sscanf Mid$(text, p), FMT_NAME, Tile_names(i), cnt
        p = p + cnt
        'increment the pointer to the start of the next record
        p = nextLine(text, p)
    Next
End Sub

Private Function mapTilesetDirToTileset(ByVal dir_ As String) As MAP_TILESET
    ' For Arizona
    If dir_ = "texpages/tertilesc1hw" Then
        mapTilesetDirToTileset = MAP_TILESET_ARIZONA
        Exit Function
    ' for Urban
    ElseIf dir_ = "texpages/tertilesc2hw" Then
        mapTilesetDirToTileset = MAP_TILESET_URBAN
        Exit Function
    ' for Rockie
    ElseIf dir_ = "texpages/tertilesc3hw" Then
        mapTilesetDirToTileset = MAP_TILESET_ROCKIES
        Exit Function
    Else
        debug_ 295, LOG_ERROR, "mapTilesetDirToTileset", "unsupported tileset: %s", dir_
        debug_ 296, LOG_POPUP, "mapTilesetDirToTileset", "This is a UNSUPPORTED map with a custom tileset.\nDefaulting to tertilesc1hw -- map may look strange!"
    End If
    mapTilesetDirToTileset = MAP_TILESET_ARIZONA ' fallback
End Function

Private Function appendToFileName(ByVal origFilename As String, ByVal appendStr As String) As String
    Dim newFilename As String, fileExtension As String, extensionSeparatorPos As Long
    newFilename = origFilename
    fileExtension = ""
    ' right-most "." index
    extensionSeparatorPos = InStrRev(origFilename, ".")
    If extensionSeparatorPos <> 0 Then
        ' separate the file extension
        fileExtension = Mid$(origFilename, extensionSeparatorPos)
        newFilename = Left$(origFilename, extensionSeparatorPos - 1)
    End If
    newFilename = newFilename & appendStr
    newFilename = newFilename & fileExtension
    appendToFileName = newFilename
End Function

Private Function getTextureVariant(ByVal origTextureFilename As String, ByVal variantStr As String) As String
    Dim variantFileName As String, variantPath As String, variantImageLoadPath As String
    variantFileName = appendToFileName(origTextureFilename, variantStr)
    variantPath = "texpages/" & variantFileName
    variantImageLoadPath = gfx_api_imageLoadFilenameFromInputFilename(variantPath)
    If PHYSFS_exists(variantImageLoadPath) <> 0 Then
        getTextureVariant = variantFileName
        Exit Function
    End If
    getTextureVariant = ""
End Function

Private Sub mapLoadTertiles(ByVal preview As Boolean, ByVal tileSet As MAP_TILESET, ByVal tertilesFile As String)
    Dim text As String, p As Long, tilename As String, textureName As String, textureType As String
    Dim textureSize As Double, numlines As Long, cnt As Long, i As Long, fileSize As Long, textureTypeIdx As Long
    ' load the override terrain types
    If Not preview And useTerrainOverrides Then
        If Not loadTerrainTypeMapOverride(tileSet) Then
            debug_ 343, LOG_POPUP, "mapLoadTertiles", "Failed to load terrain type override"
        End If
    End If
    init_tileNames tileSet
    If Not loadFileToBuffer(tertilesFile, fileLoadBuffer, FILE_LOAD_BUFFER_SIZE, fileSize) Then
        debug_ 348, LOG_FATAL, "mapLoadTertiles", "%s not found, aborting.", tertilesFile
        exit_ 3                         ' abort()
    End If
    text = fileLoadBufferText(fileSize)
    p = 1

    sscanf Mid$(text, p), FMT_NAME_COUNT, tilename, numlines, cnt
    p = p + cnt

    If InStr(tertilesFile, tilename) = 0 Then
        debug_ 357, LOG_FATAL, "mapLoadTertiles", "%s found, but was expecting %s!  Aborting.", tilename, tertilesFile
        exit_ 3                         ' abort()
    End If

    debug_ 361, LOG_TERRAIN, "mapLoadTertiles", "tilename: %s, with %d entries", tilename, numlines
    'increment the pointer to the start of the next record
    p = nextLine(text, p)
    ' groundTypes.resize(numlines): the entries that were there stay
    If numlines > 0 Then
        ReDim Preserve groundTypes(0 To numlines - 1)
    Else
        Erase groundTypes
    End If
    groundTypes_size = numlines

    For i = 0 To numlines - 1
        sscanf Mid$(text, p), FMT_TERTILE, textureType, textureName, textureSize, cnt
        p = p + cnt
        'increment the pointer to the start of the next record
        p = nextLine(text, p)

        textureTypeIdx = getTextureType(textureType)
        groundTypes(textureTypeIdx).textureName = textureName
        groundTypes(textureTypeIdx).textureSize = CSng(textureSize)
        groundTypes(textureTypeIdx).normalMapTextureName = getTextureVariant(textureName, "_nm")
        groundTypes(textureTypeIdx).specularMapTextureName = getTextureVariant(textureName, "_sm")
        groundTypes(textureTypeIdx).heightMapTextureName = getTextureVariant(textureName, "_hm")
        groundTypes(textureTypeIdx).highQualityTextures = Len(groundTypes(textureTypeIdx).normalMapTextureName) > 0 Or Len(groundTypes(textureTypeIdx).specularMapTextureName) > 0 Or Len(groundTypes(textureTypeIdx).heightMapTextureName) > 0
    Next
End Sub

Private Sub SetDecals_tileset(ByVal tileset As MAP_TILESET)
    Select Case tileset
    Case MAP_TILESET_ARIZONA
        SetDecals "tileset/arizonadecals.txt", "arizona_decals"
    ' for Urban
    Case MAP_TILESET_URBAN
        SetDecals "tileset/urbandecals.txt", "urban_decals"
    ' for Rockie
    Case MAP_TILESET_ROCKIES
        SetDecals "tileset/rockiedecals.txt", "rockie_decals"
    ' When a map uses something other than the above, we fallback to Arizona
    Case Else
        debug_ 403, LOG_ERROR, "SetDecals", "unsupported tileset: %s", tilesetDir
        ' HACK: / FIXME: For now, we just pretend this is a tertilesc1hw map.
        SetDecals "tileset/arizonadecals.txt", "arizona_decals"
    End Select
End Sub

' This is the main loading routine to get all the map's parameters set.
' Once it figures out what tileset we need, we then parse the files for that tileset.
' Currently, we only support 3 tilesets.  Arizona, Urban, and Rockie
Private Function mapLoadGroundTypes(ByVal preview As Boolean) As Boolean
    debug_ 414, LOG_TERRAIN, "mapLoadGroundTypes", "tileset: %s", tilesetDir
    Select Case currentMapTileset
    ' For Arizona
    Case MAP_TILESET_ARIZONA
        mapLoadTertiles preview, MAP_TILESET_ARIZONA, "tileset/tertilesc1hwGtype.txt"
        SetGroundForTile "tileset/arizonaground.txt", "arizona_ground"
    ' for Urban
    Case MAP_TILESET_URBAN
        mapLoadTertiles preview, MAP_TILESET_URBAN, "tileset/tertilesc2hwGtype.txt"
        SetGroundForTile "tileset/urbanground.txt", "urban_ground"
    ' for Rockie
    Case MAP_TILESET_ROCKIES
        mapLoadTertiles preview, MAP_TILESET_ROCKIES, "tileset/tertilesc3hwGtype.txt"
        SetGroundForTile "tileset/rockieground.txt", "rockie_ground"
    ' When a map uses something other than the above, we fallback to Arizona
    Case Else
        debug_ 437, LOG_ERROR, "mapLoadGroundTypes", "unsupported tileset: %s", tilesetDir
        debug_ 438, LOG_POPUP, "mapLoadGroundTypes", "This is a UNSUPPORTED map with a custom tileset.\nDefaulting to tertilesc1hw -- map may look strange!"
        ' HACK: / FIXME: For now, we just pretend this is a tertilesc1hw map.
        mapLoadTertiles preview, MAP_TILESET_ARIZONA, "tileset/tertilesc1hwGtype.txt"
        SetGroundForTile "tileset/arizonaground.txt", "arizona_ground"
    End Select

    SetDecals_tileset currentMapTileset
    mapLoadGroundTypes = True
End Function

' Parse the file to set up the ground type
Private Sub SetGroundForTile(ByVal filename As String, ByVal nametype As String)
    Dim text As String, p As Long, tilename As String, val1 As String, val2 As String, val3 As String, val4 As String
    Dim numlines As Long, cnt As Long, i As Long, fileSize As Long

    If Not loadFileToBuffer(filename, fileLoadBuffer, FILE_LOAD_BUFFER_SIZE, fileSize) Then
        debug_ 460, LOG_FATAL, "SetGroundForTile", "%s not found, aborting.", filename
        exit_ 3                         ' abort()
    End If
    text = fileLoadBufferText(fileSize)
    p = 1

    sscanf Mid$(text, p), FMT_NAME_COUNT, tilename, numlines, cnt
    p = p + cnt

    If tilename <> nametype Then
        debug_ 469, LOG_FATAL, "SetGroundForTile", "%s found, but was expecting %s, aborting.", tilename, nametype
        exit_ 3                         ' abort()
    End If

    debug_ 473, LOG_TERRAIN, "SetGroundForTile", "tilename: %s, with %d entries", tilename, numlines
    'increment the pointer to the start of the next record
    p = nextLine(text, p)

    numTile_types = numlines
    If numlines > 0 Then ReDim map_(0 To numlines * 2 * 2 - 1) Else Erase map_

    For i = 0 To numlines - 1
        sscanf Mid$(text, p), FMT_GROUND, val1, val2, val3, val4, cnt
        p = p + cnt
        'increment the pointer to the start of the next record
        p = nextLine(text, p)

        ' inline int iA(int i, int j, int k){ return i*N2*N3 + j*N3 + k; }
        ' in case it isn't obvious, this is a 3D array, and using pointer math to access each element.
        ' so map[10][0][1] would be map[10*2*2 + 0 + 1] == map[41]
        ' map[10][1][0] == map[10*2*2 + 2 + 0] == map[42]
        map_(i * 2 * 2 + 0 * 2 + 0) = getTextureType(val4)
        map_(i * 2 * 2 + 0 * 2 + 1) = getTextureType(val2)
        map_(i * 2 * 2 + 1 * 2 + 0) = getTextureType(val3)
        map_(i * 2 * 2 + 1 * 2 + 1) = getTextureType(val1)
    Next
End Sub

' getTextureType() -- just returns the value for that texture type.
Private Function getTextureType(ByVal textureType As String) As Long
    Dim i As Long
    For i = 0 To numTile_names - 1
        If textureType = Tile_names(i) Then
            getTextureType = i
            Exit Function
        End If
    Next
    debug_ 509, LOG_FATAL, "getTextureType", "unknown type [%s] found, aborting!", textureType
    exit_ 3                             ' abort()
End Function

' groundFromMapTile() just a simple lookup table, using pointers to access the 3D map array
'	(quasi) pointer math is: map[num elements][2][2]
'	so map[10][0][1] would be map[10*2*2 + 0*2 + 1] == map[41]
Private Function groundFromMapTile(ByVal tile As Long, ByVal j As Long, ByVal k As Long) As Long
    Dim tileNumber As Long
    tileNumber = TileNumber_tile(tile)
    If tileNumber >= numTile_types Then
        debug_ 521, LOG_INFO, "groundFromMapTile", "Invalid ground tile number: %d", tileNumber
        groundFromMapTile = 0
        Exit Function
    End If
    groundFromMapTile = map_(TileNumber_tile(tile) * 2 * 2 + j * 2 + k)
End Function

Private Sub rotFlip(ByVal tile As Long, i As Long, j As Long)
    Dim texture As Long, rot As Long
    Dim tmpMap(0 To 1, 0 To 1) As Long, invmap(0 To 3, 0 To 1) As Long

    texture = TileNumber_texture(tile)
    If (texture And TILE_XFLIP) <> 0 Then
        i = 1 - i
    End If
    If (texture And TILE_YFLIP) <> 0 Then
        j = 1 - j
    End If

    tmpMap(0, 0) = 0: invmap(0, 0) = 0: invmap(0, 1) = 0
    tmpMap(1, 0) = 1: invmap(1, 0) = 1: invmap(1, 1) = 0
    tmpMap(1, 1) = 2: invmap(2, 0) = 1: invmap(2, 1) = 1
    tmpMap(0, 1) = 3: invmap(3, 0) = 0: invmap(3, 1) = 1
    rot = tmpMap(i, j)
    rot = rot - (texture And TILE_ROTMASK) \ BIT_(TILE_ROTSHIFT)
    Do While rot < 0
        rot = rot + 4
    Loop
    i = invmap(rot, 0)
    j = invmap(rot, 1)
End Sub

' Tries to figure out what ground type a grid point is from the surrounding tiles
Private Function determineGroundType(ByVal x As Long, ByVal y As Long, ByVal tileset As String) As Long
    Dim ground_(0 To 1, 0 To 1) As Long, votes(0 To 1, 0 To 1) As Long, weight(0 To 1, 0 To 1) As Long
    Dim i As Long, j As Long, tile As Long, a As Long, b As Long, best As Long, psTile As Long

    If x < 0 Or y < 0 Or x >= mapWidth Or y >= mapHeight Then
        determineGroundType = 0 ' just return the first ground type
        Exit Function
    End If

    ' check what tiles surround this grid point
    For i = 0 To 1
        For j = 0 To 1
            If x + i - 1 < 0 Or y + j - 1 < 0 Or x + i - 1 >= mapWidth Or y + j - 1 >= mapHeight Then
                psTile = -1                 ' nullptr
                tile = 0
            Else
                psTile = mapTile(x + i - 1, y + j - 1)
                tile = psMapTiles(psTile).texture
            End If
            a = i
            b = j
            rotFlip tile, a, b
            ground_(i, j) = groundFromMapTile(tile, a, b)

            votes(i, j) = 0
            ' votes are weighted, some tiles have more weight than others
            weight(i, j) = 10

            If psTile <> -1 Then
                ' cliff tiles have higher priority, to be clearly visible
                If terrainType(psTile) = TER_CLIFFFACE Then
                    weight(i, j) = 100
                End If
                ' water bottom has lower priority, to stay inside water
                If terrainType(psTile) = TER_WATER Then
                    weight(i, j) = 1
                End If
            End If
        Next
    Next

    ' now vote, because some maps have seams
    For i = 0 To 1
        For j = 0 To 1
            For a = 0 To 1
                For b = 0 To 1
                    If ground_(i, j) = ground_(a, b) Then
                        votes(i, j) = votes(i, j) + weight(a, b)
                    End If
                Next
            Next
        Next
    Next
    ' and determine the winner
    best = -1
    a = 0
    b = 0
    For i = 0 To 1
        For j = 0 To 1
            If votes(i, j) > best Or (votes(i, j) = best And ground_(i, j) < ground_(a, b)) Then
                best = votes(i, j)
                a = i
                b = j
            End If
        Next
    Next
    determineGroundType = ground_(a, b)
End Function


' SetDecals()
' reads in the decal array for the requested tileset.
Private Sub SetDecals(ByVal filename As String, ByVal decal_type As String)
    Dim text As String, p As Long, decalname As String, numlines As Long, cnt As Long, i As Long, tiledecal As Long
    Dim fileSize As Long

    If Not loadFileToBuffer(filename, fileLoadBuffer, FILE_LOAD_BUFFER_SIZE, fileSize) Then
        debug_ 660, LOG_POPUP, "SetDecals", "%s not found, aborting.", filename
        exit_ 3                         ' abort()
    End If
    text = fileLoadBufferText(fileSize)
    p = 1

    sscanf Mid$(text, p), FMT_NAME_COUNT, decalname, numlines, cnt
    p = p + cnt

    If decalname <> decal_type Then
        debug_ 669, LOG_POPUP, "SetDecals", "%s found, but was expecting %s, aborting.", decalname, decal_type
        exit_ 3                         ' abort()
    End If

    debug_ 673, LOG_TERRAIN, "SetDecals", "reading: %s, with %d entries", filename, numlines
    'increment the pointer to the start of the next record
    p = nextLine(text, p)
    ' value initialization sets everything to false.
    ' (VB) one more entry: the range check below lets MAX_TERRAIN_TILES itself through
    ReDim mapDecals(0 To MAX_TERRAIN_TILES)
    mapDecals_allocated = True

    For i = 0 To numlines - 1
        tiledecal = -1
        sscanf Mid$(text, p), "%d%n", tiledecal, cnt
        p = p + cnt
        'increment the pointer to the start of the next record
        p = nextLine(text, p)
        If ULT(MAX_TERRAIN_TILES, tiledecal) Then
            debug_ 688, LOG_ERROR, "SetDecals", "Tile index is out of range!  Was %d, our max is %d", tiledecal, MAX_TERRAIN_TILES
            GoTo continue_
        End If
        mapDecals(tiledecal) = True
continue_:
    Next
End Sub
' hasDecals()
' Checks to see if the requested tile has a decal on it or not.
Private Function hasDecals(ByVal i As Long, ByVal j As Long) As Boolean
    Dim index As Long
    index = 0
    index = TileNumber_tile(psMapTiles(mapTile(i, j)).texture)
    If index > MAX_TERRAIN_TILES Then
        debug_ 702, LOG_FATAL, "hasDecals", "Tile index is out of range!  Was %d, our max is %d", index, MAX_TERRAIN_TILES
        exit_ 3                         ' abort()
    End If
    hasDecals = mapDecals(index)
End Function
' mapSetGroundTypes()
' Sets the ground type to be a decal or not
Private Function mapSetGroundTypes() As Boolean
    Dim i As Long, j As Long, psTile As Long
    For j = 0 To mapHeight - 1
        For i = 0 To mapWidth - 1
            psTile = mapTile(i, j)

            psMapTiles(psTile).ground = determineGroundType(i, j, tilesetDir) And &HFF&

            If hasDecals(i, j) Then
                SET_TILE_DECAL psTile
            Else
                CLEAR_TILE_DECAL psTile
            End If
        Next
    Next
    mapSetGroundTypes = True
End Function

Public Function mapReloadGroundTypes() As Boolean
    If StrPtr(tilesetDir) = 0 Then
        mapReloadGroundTypes = False
        Exit Function
    End If
    mapLoadGroundTypes False
    If Not mapSetGroundTypes() Then
        mapReloadGroundTypes = False
        Exit Function
    End If
    mapReloadGroundTypes = True
End Function

Private Function isWaterVertex(ByVal x As Long, ByVal y As Long) As Boolean
    If x < 1 Or y < 1 Or x > mapWidth - 1 Or y > mapHeight - 1 Then
        isWaterVertex = False
        Exit Function
    End If
    isWaterVertex = terrainType(mapTile(x, y)) = TER_WATER And terrainType(mapTile(x - 1, y)) = TER_WATER _
                    And terrainType(mapTile(x, y - 1)) = TER_WATER And terrainType(mapTile(x - 1, y - 1)) = TER_WATER
End Function

Private Sub generateRiverbed()
    Dim mt As MersenneTwister, maxIdx As Long, idx() As Long, x As Long, y As Long, l As Long, val_ As Long
    Dim rowOffset As Long, k As Long, u As Double, m As Double, t As Long
    MersenneTwister_construct mt, 12345     ' 12345 = random seed.
    maxIdx = 1
    If Not (mapWidth > 0 And mapHeight > 0) Then ASSERT_FAILURE "mapWidth > 0 && mapHeight > 0", "src/map.cpp:760", "generateRiverbed", "Invalid map width or height (%d x %d)", mapWidth, mapHeight: Exit Sub
    ReDim idx(0 To mapWidth * mapHeight - 1)
    l = 0

    For y = 0 To mapHeight - 1
        For x = 0 To mapWidth - 1
            ' initially set the seabed index to 0 for ground and 100 for water
            val_ = 100 * -CLng(isWaterVertex(x, y))
            idx(x + (y * mapWidth)) = val_
            If val_ > 0 Then
                l = l + 1
            End If
        Next
    Next
    debug_ 777, LOG_TERRAIN, "generateRiverbed", "Generating riverbed for %d water tiles.", l
    If l = 0 Then ' no water on map
        Exit Sub
    End If
    l = 0
    Do
        maxIdx = 1
        For y = 1 To mapHeight - 3
            For x = 1 To mapWidth - 3
                rowOffset = (y * mapWidth)
                k = x + rowOffset
                If idx(k) > 0 Then
                    idx(k) = (idx((x - 1) + rowOffset) + idx(x + ((y - 1) * mapWidth)) + idx(x + ((y + 1) * mapWidth)) + idx((x + 1) + rowOffset)) \ 4
                    If idx(k) > maxIdx Then
                        maxIdx = idx(k)
                    End If
                End If
            Next
        Next
        l = l + 1
        debug_ 803, LOG_TERRAIN, "generateRiverbed", "%d%% completed after %d iterations", 10 * (100 - maxIdx), l
    Loop While maxIdx > 90 And l < 20

    For y = 0 To mapHeight - 1
        For x = 0 To mapWidth - 1
            k = x + (y * mapWidth)
            If idx(k) > maxIdx Then
                idx(k) = maxIdx
            End If
            If idx(k) < 1 Then
                idx(k) = 1
            End If
            If isWaterVertex(x, y) Then
                ' (the uint32_t remainder, then int arithmetic: no value here leaves int's range)
                u = U32(MersenneTwister_u32(mt))
                m = maxIdx \ 6 + 1
                l = (WATER_MAX_DEPTH + 1 - WATER_MIN_DEPTH) * (maxIdx - idx(k) - CLng(u - Int(u / m) * m))
                t = mapTile(x, y)
                psMapTiles(t).height = psMapTiles(t).height - (WATER_MIN_DEPTH - (l \ maxIdx))
            End If
        Next
    Next

End Sub

' Initialise the map structure
Public Function mapLoad(ByVal filename As String) As Boolean
    Dim mapIO As WzMapPhysFSIO, debugLoggerInstance As WzMapDebugLogger, loadedMap As MapData
    Set mapIO = New WzMapPhysFSIO
    Set debugLoggerInstance = New WzMapDebugLogger

    Set loadedMap = loadMapData(filename, mapIO, debugLoggerInstance)
    If loadedMap Is Nothing Then
        ' loadMapData call handles logging errors
        mapLoad = False
        Exit Function
    End If
    mapLoad = mapLoadFromWzMapData(loadedMap)

End Function

' -----------------------------------------------------------------------------------------
' load up a terrain tile type map file
Public Function loadTerrainTypeMap(ByVal ttypeData As TerrainTypeData) As Boolean
    Dim quantity As Long, i As Long, typ As Long
    If Not (Not ttypeData Is Nothing) Then ASSERT_FAILURE "ttypeData != nullptr", "src/map.cpp:1031", "loadTerrainTypeMap", "No terrain type data": loadTerrainTypeMap = False: Exit Function

    ' reset the terrain table
    For i = 0 To MAX_TILE_TEXTURES - 1
        terrainTypes(i) = 0
    Next

    quantity = ttypeData.terrainTypes_size()
    If quantity >= MAX_TILE_TEXTURES Then
        ' Workaround for fugly map editor bug, since we can't fix the map editor
        quantity = MAX_TILE_TEXTURES - 1
    End If
    For i = 0 To quantity - 1
        typ = ttypeData.terrainTypes(i)
        If typ > TER_MAX Then
            debug_ 1047, LOG_ERROR, "loadTerrainTypeMap", "loadTerrainTypeMap: terrain type out of range"
            loadTerrainTypeMap = False
            Exit Function
        End If

        terrainTypes(i) = typ And &HFF&
    Next

    loadTerrainTypeMap = True
End Function

' Initialise the map structure
Public Function mapLoadFromWzMapData(ByVal loadedMap As MapData) As Boolean
    Dim width As Long, height As Long, i As Long, k As Long, gwIdx As Long, h As Long
    Const preview As Boolean = False

    width = loadedMap.width
    height = loadedMap.height

    ' See if this is the first time a map has been loaded
    If Not (Not psMapTiles_allocated) Then ASSERT_FAILURE "psMapTiles == nullptr", "src/map.cpp:1067", "mapLoadFromWzMapData", "Map has not been cleared before calling mapLoad()!"

    ' Allocate the memory for the map
    If width * height > 0 Then
        ReDim psMapTiles(0 To width * height - 1)
    Else
        Erase psMapTiles
    End If
    psMapTiles_allocated = True
    LightMap_reset width, height

    mapWidth = width
    mapHeight = height

    ' FIXME: the map preview code loads the map without setting the tileset
    If StrPtr(tilesetDir) = 0 Then
        tilesetDir = "texpages/tertilesc1hw"
    End If

    currentMapTileset = mapTilesetDirToTileset(tilesetDir)

    ' load the ground types
    If Not mapLoadGroundTypes(preview) Then
        mapLoadFromWzMapData = False
        Exit Function
    End If

    If Not preview Then
        'preload the terrain textures
        loadTerrainTextures currentMapTileset
    End If

    'load in the map data itself

    ' Load in the map data
    For i = 0 To mapWidth * mapHeight - 1
        h = loadedMap.mMapTiles_height(i)
        If Not (h <= TILE_MAX_HEIGHT) Then ASSERT_FAILURE "loadedMap->mMapTiles[i].height <= (255 * 2)", "src/map.cpp:1102", "mapLoadFromWzMapData", "Tile height (%u) exceeds TILE_MAX_HEIGHT (%zu)", h, CLng(TILE_MAX_HEIGHT)
        psMapTiles(i).texture = loadedMap.mMapTiles_texture(i)
        psMapTiles(i).height = h

        ' Visibility stuff
        For k = 0 To MAX_PLAYERS - 1
            psMapTiles(i).watchers(k) = 0
            psMapTiles(i).sensors(k) = 0
            psMapTiles(i).jammers(k) = 0
        Next
        psMapTiles(i).sensorBits = 0
        psMapTiles(i).jammerBits = 0
        psMapTiles(i).tileExploredBits = 0
    Next

    If preview Then
        ' no need to do anything else for the map preview
        mapLoadFromWzMapData = True
        Exit Function
    End If

    gwIdx = 0
    For gwIdx = 0 To loadedMap.mGateways_size() - 1
        If Not gwNewGateway(loadedMap.mGateways_x1(gwIdx), loadedMap.mGateways_y1(gwIdx), loadedMap.mGateways_x2(gwIdx), loadedMap.mGateways_y2(gwIdx)) Then
            debug_ 1126, LOG_ERROR, "mapLoadFromWzMapData", "Unable to add gateway %zu - dropping it", gwIdx
        End If
    Next

    If Not afterMapLoad() Then
        mapLoadFromWzMapData = False
        Exit Function
    End If

    mapLoadFromWzMapData = True
End Function

Private Function afterMapLoad() As Boolean
    Dim x As Long, y As Long, mapSize As Long, psTile As Long, t As Long
    If Not mapSetGroundTypes() Then
        afterMapLoad = False
        Exit Function
    End If

    For y = 0 To mapHeight - 1
        For x = 0 To mapWidth - 1
            ' FIXME: magic number
            t = mapTile(x, y)
            psMapTiles(t).waterLevel = psMapTiles(t).height - world_coord(1) \ 3
        Next
    Next
    generateRiverbed

    ' set up the scroll mins and maxs - set values to valid ones for any new map
    scrollMinX = 0
    scrollMinY = 0
    scrollMaxX = mapWidth
    scrollMaxY = mapHeight

    ' Allocate aux maps
    If Not (mapWidth >= 0 And mapHeight >= 0) Then ASSERT_FAILURE "mapWidth >= 0 && mapHeight >= 0", "src/map.cpp:1162", "afterMapLoad", "Invalid mapWidth or mapHeight (%d x %d)", mapWidth, mapHeight
    mapSize = mapWidth * mapHeight
    If mapSize > 0 Then
        ReDim psBlockMap(0 To AUX_MAX - 1, 0 To mapSize - 1)
        ReDim psAuxMap(0 To MAX_PLAYERS + AUX_MAX - 1, 0 To mapSize - 1)
    End If
    auxMaps_allocated = True

    ' Set our blocking bits
    For y = 0 To mapHeight - 1
        For x = 0 To mapWidth - 1
            psTile = mapTile(x, y)

            auxClearBlocking x, y, AUXBITS_ALL
            auxClearAll x, y, AUXBITS_ALL

            ' All tiles outside of the map and on map border are blocking.
            If x < 1 Or y < 1 Or x > mapWidth - 1 Or y > mapHeight - 1 Then
                auxSetBlocking x, y, AUXBITS_ALL ' block everything
            End If
            If terrainType(psTile) = TER_WATER Then
                auxSetBlocking x, y, WATER_BLOCKED
            Else
                auxSetBlocking x, y, LAND_BLOCKED
            End If
            If terrainType(psTile) = TER_CLIFFFACE Then
                auxSetBlocking x, y, FEATURE_BLOCKED
            End If
        Next
    Next

    ' Set continents. This should ideally be done in advance by the map editor.
    mapFloodFillContinents

    afterMapLoad = True
End Function

' Shutdown the map module
Public Function mapShutdown() As Boolean
    ' (no danger thread to stop: see the module's comment)

    Erase floodbucket_x
    Erase floodbucket_y
    Erase mapDecals
    mapDecals_allocated = False
    Erase psBlockMap
    Erase psAuxMap
    auxMaps_allocated = False

    Erase map_
    Erase groundTypes
    groundTypes_size = 0
    Erase psMapTiles
    psMapTiles_allocated = False
    mapWidth = 0
    mapHeight = 0
    numTile_names = 0
    Erase Tile_names
    If StrPtr(tilesetDir) <> 0 Then
        tilesetDir = vbNullString
    End If
    mapShutdown = True
End Function

' Intersect a tile with a line and report the points of intersection
' line is gives as point plus 2d directional vector
' returned are two coordinates at the edge
' true if the intersection also crosses the tile split line
' (which has to be taken into account)
Public Function map_Intersect(Cx As Long, Cy As Long, Vx As Long, Vy As Long, Sx As Long, Sy As Long) As Boolean
    Dim x As Long, y As Long, ox As Long, oy As Long, Dx As Long, Dy As Long, tileX As Long, tileY As Long
    Dim ily As Long, iry As Long, itx As Long, ibx As Long

    ' dereference pointers
    x = Cx
    y = Cy
    Dx = Vx
    Dy = Vy

    ' Turn into tile coordinates
    tileX = map_coord(x)
    tileY = map_coord(y)

    ' Inter tile comp
    ox = map_round(x)
    oy = map_round(y)

    ' allow backwards tracing
    If ox = 0 And Dx < 0 Then
        tileX = tileX - 1
        ox = TILE_UNITS
    End If
    If oy = 0 And Dy < 0 Then
        tileY = tileY - 1
        oy = TILE_UNITS
    End If

    Cx = -4 * TILE_UNITS ' to trigger assertion
    Cy = -4 * TILE_UNITS
    Vx = -4 * TILE_UNITS
    Vy = -4 * TILE_UNITS

    ' calculate intersection point on the left and right (if any)
    ily = y - 4 * TILE_UNITS ' make sure initial value is way outside of tile
    iry = y - 4 * TILE_UNITS
    If Dx <> 0 Then
        ily = y - ox * Dy \ Dx
        iry = y + (TILE_UNITS - ox) * Dy \ Dx
    End If
    ' calculate intersection point on top and bottom (if any)
    itx = x - 4 * TILE_UNITS ' make sure initial value is way outside of tile
    ibx = x - 4 * TILE_UNITS
    If Dy <> 0 Then
        itx = x - oy * Dx \ Dy
        ibx = x + (TILE_UNITS - oy) * Dx \ Dy
    End If

    ' line comes from the left?
    If Dx >= 0 Then
        If map_coord(ily) = tileY Or map_coord(ily - 1) = tileY Then
            Cx = world_coord(tileX)
            Cy = ily
        End If
        If map_coord(iry) = tileY Or map_coord(iry - 1) = tileY Then
            Vx = world_coord(tileX + 1)
            Vy = iry
        End If
    Else
        If map_coord(ily) = tileY Or map_coord(ily - 1) = tileY Then
            Vx = world_coord(tileX)
            Vy = ily
        End If
        If map_coord(iry) = tileY Or map_coord(iry - 1) = tileY Then
            Cx = world_coord(tileX + 1)
            Cy = iry
        End If
    End If
    ' line comes from the top?
    If Dy >= 0 Then
        If map_coord(itx) = tileX Or map_coord(itx - 1) = tileX Then
            Cx = itx
            Cy = world_coord(tileY)
        End If
        If map_coord(ibx) = tileX Or map_coord(ibx - 1) = tileX Then
            Vx = ibx
            Vy = world_coord(tileY + 1)
        End If
    Else
        If map_coord(itx) = tileX Or map_coord(itx - 1) = tileX Then
            Vx = itx
            Vy = world_coord(tileY)
        End If
        If map_coord(ibx) = tileX Or map_coord(ibx - 1) = tileX Then
            Cx = ibx
            Cy = world_coord(tileY + 1)
        End If
    End If
    ' assertions, no intersections outside of tile
    If Not (Cx >= world_coord(tileX) And Cx <= world_coord(tileX + 1)) Then ASSERT_FAILURE "*Cx >= world_coord(tileX) && *Cx <= world_coord(tileX + 1)", "src/map.cpp:1412", "map_Intersect", "map_Intersect(): tile Bounds %i %i, %i %i -> %i,%i,%i,%i", x, y, Dx, Dy, Cx, Cy, Vx, Vy
    If Not (Cy >= world_coord(tileY) And Cy <= world_coord(tileY + 1)) Then ASSERT_FAILURE "*Cy >= world_coord(tileY) && *Cy <= world_coord(tileY + 1)", "src/map.cpp:1413", "map_Intersect", "map_Intersect(): tile Bounds %i %i, %i %i -> %i,%i,%i,%i", x, y, Dx, Dy, Cx, Cy, Vx, Vy
    If Not (Vx >= world_coord(tileX) And Vx <= world_coord(tileX + 1)) Then ASSERT_FAILURE "*Vx >= world_coord(tileX) && *Vx <= world_coord(tileX + 1)", "src/map.cpp:1414", "map_Intersect", "map_Intersect(): tile Bounds %i %i, %i %i -> %i,%i,%i,%i", x, y, Dx, Dy, Cx, Cy, Vx, Vy
    If Not (Vy >= world_coord(tileY) And Vy <= world_coord(tileY + 1)) Then ASSERT_FAILURE "*Vy >= world_coord(tileY) && *Vy <= world_coord(tileY + 1)", "src/map.cpp:1415", "map_Intersect", "map_Intersect(): tile Bounds %i %i, %i %i -> %i,%i,%i,%i", x, y, Dx, Dy, Cx, Cy, Vx, Vy
    If Not (tileX >= 0 And tileY >= 0 And tileX < mapWidth And tileY < mapHeight) Then ASSERT_FAILURE "tileX >= 0 && tileY >= 0 && tileX < mapWidth && tileY < mapHeight", "src/map.cpp:1416", "map_Intersect", "map_Intersect(): map Bounds %i %i, %i %i -> %i,%i,%i,%i", x, y, Dx, Dy, Cx, Cy, Vx, Vy

    'calculate midway line intersection points
    If ((map_coord(itx) = tileX) = (map_coord(ily) = tileY)) And ((map_coord(ibx) = tileX) = (map_coord(iry) = tileY)) Then
        ' line crosses diagonal only
        If Dx - Dy = 0 Then
            map_Intersect = False
            Exit Function
        End If
        Sx = world_coord(tileX) + (Dx * oy - Dy * ox) \ (Dx - Dy)
        Sy = world_coord(tileY) + (Dx * oy - Dy * ox) \ (Dx - Dy)
        If map_coord(Sx) <> tileX Or map_coord(Sy) <> tileY Then
            map_Intersect = False
            Exit Function
        End If
        map_Intersect = True
        Exit Function
    ElseIf ((map_coord(ibx) = tileX) = (map_coord(ily) = tileY)) And ((map_coord(itx) = tileX) = (map_coord(iry) = tileY)) Then
        'line crosses anti-diagonal only
        If Dx + Dy = 0 Then
            map_Intersect = False
            Exit Function
        End If
        Sx = world_coord(tileX) + (Dx * (TILE_UNITS - oy) + Dy * ox) \ (Dx + Dy)
        Sy = world_coord(tileY) + (Dy * (TILE_UNITS - ox) + Dx * oy) \ (Dx + Dy)
        If map_coord(Sx) <> tileX Or map_coord(Sy) <> tileY Then
            map_Intersect = False
            Exit Function
        End If
        map_Intersect = True
        Exit Function
    Else
        'line crosses both tile diagonals
        'TODO: trunk divides tiles into 4 parts instead of 2 in 2.3.
        'We would need to check and return both intersections here now,
        'but that would require an additional return parameter!
        'Instead we check only one of them and know it might be wrong!
        If Dx + Dy <> 0 Then
            ' check anti-diagonal
            Sx = world_coord(tileX) + (Dx * (TILE_UNITS - oy) + Dy * ox) \ (Dx + Dy)
            Sy = world_coord(tileY) + (Dy * (TILE_UNITS - ox) + Dx * oy) \ (Dx + Dy)
            If map_coord(Sx) = tileX And map_coord(Sy) = tileY Then
                map_Intersect = True
                Exit Function
            End If
        End If
        If Dx - Dy <> 0 Then
            ' check diagonal
            Sx = world_coord(tileX) + (Dx * oy - Dy * ox) \ (Dx - Dy)
            Sy = world_coord(tileY) + (Dx * oy - Dy * ox) \ (Dx - Dy)
            If map_coord(Sx) = tileX And map_coord(Sy) = tileY Then
                map_Intersect = True
                Exit Function
            End If
        End If
    End If

    map_Intersect = False
End Function

' Rotate vector clockwise by quadrant*90 degrees around (TILE_UNITS/2, TILE_UNITS/2). (Considering x to be to the right, and y down.)
Private Sub rotateWorldQuadrant(x As Long, y As Long, ByVal quadrant As Long)
    Dim t As Long
    Select Case quadrant And 3
    Case 0                                  ' 0 degrees.
    Case 1: t = x: x = TILE_UNITS - y: y = t                ' 90 degrees clockwise.
    Case 2: x = TILE_UNITS - x: y = TILE_UNITS - y          ' 180 degrees.
    Case 3: t = x: x = y: y = TILE_UNITS - t                ' 90 degrees anticlockwise.
    End Select
End Sub

' Returns (0, 0) rotated clockwise quadrant*90 degrees around (1/2, 1/2). (Considering x to be to the right, and y down.)
Private Sub quadrantCorner(ByVal quadrant As Long, dx As Long, dy As Long)
    dx = Choose((quadrant And 3) + 1, 0, 1, 1, 0)
    dy = Choose((quadrant And 3) + 1, 0, 0, 1, 1)
End Sub

' Returns (0, -1) rotated clockwise quadrant*90 degrees around (0, 0). (Considering x to be to the right, and y down.)
Private Sub quadrantDelta(ByVal quadrant As Long, dx As Long, dy As Long)
    dx = Choose((quadrant And 3) + 1, 0, 1, 0, -1)
    dy = Choose((quadrant And 3) + 1, -1, 0, 1, 0)
End Sub


Private Function fracTest(ByVal numerA As Long, ByVal denomA As Long, ByVal numerB As Long, ByVal denomB As Long) As Boolean
    ' (int64_t products as Decimals)
    If denomA > 0 And numerA >= 0 Then
        If denomB <= 0 Or numerB < 0 Then
            fracTest = True
        Else
            fracTest = CDec(numerA) * denomB < CDec(numerB) * denomA
        End If
    End If
End Function

' the line's src and dst are (srcX, srcY, srcZ) and (dstX, dstY, dstZ); UINT32_MAX (-1) when there is no intersection
Public Function map_LineIntersect(ByVal srcX As Long, ByVal srcY As Long, ByVal srcZ As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstZ As Long, ByVal tMax As Long) As Long
    Dim tileX As Long, tileY As Long, quadrant As Long, height(0 To 3) As Long, q As Long, cx As Long, cy As Long
    Dim difX As Long, difY As Long, difZ As Long, numer(0 To 4) As Long, denom(0 To 4) As Long
    Dim normalX As Long, normalY As Long, normalZ As Long, firstIntersection As Long, test As Long, dx As Long, dy As Long
    ' Transform src and dst to a coordinate system such that the tile quadrant containing src has
    ' corners at (0, 0), (TILE_UNITS, 0), (TILE_UNITS/2, TILE_UNITS/2).
    tileX = map_coord(srcX)
    tileY = map_coord(srcY)
    srcX = srcX - world_coord(tileX): srcY = srcY - world_coord(tileY)
    dstX = dstX - world_coord(tileX): dstY = dstY - world_coord(tileY)
    '            +0+
    ' quadrant = 3x1
    '            +2+
    quadrant = (-CLng(srcX < srcY) * 3) Xor -CLng(TILE_UNITS - srcX < srcY)
    rotateWorldQuadrant srcX, srcY, -quadrant
    rotateWorldQuadrant dstX, dstY, -quadrant
    Do
        For q = 0 To 3
            quadrantCorner quadrant + q, cx, cy
            height(q) = map_TileHeightSurface(tileX + cx, tileY + cy)
        Next
        difX = dstX - srcX: difY = dstY - srcY: difZ = dstZ - srcZ
        '     We are considering the volume of a quadrant (the volume above a quarter of a map tile, which is
        ' a degenerate tetrahedron with a point at infinity). We have a line segment, and want to know where
        ' it exits the quadrant volume.
        '     There are 5 possible cases. Cases 0-2, our line can exit one of the three sides of the quadrant
        ' volume (and pass into a neighbouring quadrant volume), or case 3, can exit through the bottom of the
        ' quadrant volume (which means intersecting the terrain), or case 4, the line segment can end (which
        ' means reaching the destination with no intersection.
        '     Note that the height of the centre of the tile is the average of the corners, such that a tile
        ' consists of 4 flat triangles (which are not (in general) parallel to each other).
        ' +--0--+
        '  \ 3 /
        '   2 1
        '    +
        ' Denominators are positive iff we are going in the direction of the line. First line crossed has the smallest fraction.
        ' numer/denom gives the intersection times for the 5 cases.
        numer(0) = -(-srcY)
        denom(0) = -difY
        numer(1) = TILE_UNITS - (srcX + srcY)
        denom(1) = difX + difY
        numer(2) = -(-srcX + srcY)
        denom(2) = -difX + difY
        normalX = 2 * (height(1) - height(0)): normalY = height(2) + height(3) - height(0) - height(1): normalZ = -2 * TILE_UNITS ' Normal pointing down, and not normalised.
        numer(3) = height(0) * normalZ - (srcX * normalX + srcY * normalY + srcZ * normalZ)
        denom(3) = (difX * normalX + difY * normalY + difZ * normalZ)
        numer(4) = 1
        denom(4) = 1
        firstIntersection = 0
        For test = 0 To 4
            If Not fracTest(numer(firstIntersection), denom(firstIntersection), numer(test), denom(test)) Then
                firstIntersection = test
            End If
        Next
        Select Case firstIntersection
        Case 0  ' Cross top line first (the tile boundary).
            quadrantDelta quadrant, dx, dy
            tileX = tileX + dx: tileY = tileY + dy
            quadrant = quadrant + 2
            rotateWorldQuadrant srcX, srcY, -2: srcY = srcY - TILE_UNITS
            rotateWorldQuadrant dstX, dstY, -2: dstY = dstY - TILE_UNITS

            If tileX < 0 Or tileX >= mapWidth Or tileY < 0 Or tileY >= mapHeight Then
                ' Intersect edge of map.
                map_LineIntersect = FromU32Dec(Fix(CDec(U32(tMax)) * numer(firstIntersection) / denom(firstIntersection)))
                Exit Function
            End If
        Case 1  ' Cross bottom-right line first.
            ' Change to the new quadrant, and transform appropriately.
            quadrant = quadrant + 1
            rotateWorldQuadrant srcX, srcY, -1
            rotateWorldQuadrant dstX, dstY, -1
        Case 2  ' Cross bottom-left line first.
            ' Change to the new quadrant, and transform appropriately.
            quadrant = quadrant - 1
            rotateWorldQuadrant srcX, srcY, 1
            rotateWorldQuadrant dstX, dstY, 1
        Case 3  ' Intersect terrain!
            map_LineIntersect = FromU32Dec(Fix(CDec(U32(tMax)) * numer(firstIntersection) / denom(firstIntersection)))
            Exit Function
        Case 4  ' Line segment ends.
            map_LineIntersect = &HFFFFFFFF
            Exit Function
        End Select
    Loop
End Function

' The max height of the terrain and water at the specified world coordinates
Public Function map_Height(ByVal x As Long, ByVal y As Long) As Long
    Dim tileX As Long, tileY As Long, i As Long, j As Long, height(0 To 1, 0 To 1) As Long, center As Long
    Dim onTileX As Long, onTileY As Long, left_ As Long, right_ As Long, middle As Long
    Dim onBottom As Long, result As Long, towardsCenter As Long, towardsRight As Long

    ' Clamp x and y values to actual ones
    ' Give one tile worth of leeway before asserting, for units/transporters coming in from off-map.
    If Not (x >= -TILE_UNITS) Then ASSERT_FAILURE "x >= -(1<<7)", "src/map.cpp:1620", "map_Height", "map_Height: x value is too small (%d,%d) in %dx%d", map_coord(x), map_coord(y), mapWidth, mapHeight
    If Not (y >= -TILE_UNITS) Then ASSERT_FAILURE "y >= -(1<<7)", "src/map.cpp:1621", "map_Height", "map_Height: y value is too small (%d,%d) in %dx%d", map_coord(x), map_coord(y), mapWidth, mapHeight
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    If Not (x < world_coord(mapWidth) + TILE_UNITS) Then ASSERT_FAILURE "x < world_coord(mapWidth) + (1<<7)", "src/map.cpp:1624", "map_Height", "map_Height: x value is too big (%d,%d) in %dx%d", map_coord(x), map_coord(y), mapWidth, mapHeight
    If Not (y < world_coord(mapHeight) + TILE_UNITS) Then ASSERT_FAILURE "y < world_coord(mapHeight) + (1<<7)", "src/map.cpp:1625", "map_Height", "map_Height: y value is too big (%d,%d) in %dx%d", map_coord(x), map_coord(y), mapWidth, mapHeight
    If world_coord(mapWidth) - 1 < x Then x = world_coord(mapWidth) - 1
    If world_coord(mapHeight) - 1 < y Then y = world_coord(mapHeight) - 1

    ' on which tile are these coords?
    tileX = map_coord(x)
    tileY = map_coord(y)

    ' where on the tile? (scale to (0,1))
    onTileX = x - world_coord(tileX)
    onTileY = y - world_coord(tileY)

    ' get the height for the corners and center
    center = 0
    For i = 0 To 1
        For j = 0 To 1
            height(i, j) = map_TileHeightSurface(tileX + i, tileY + j)
            center = center + height(i, j)
        Next
    Next
    center = center \ 4

    ' we have:
    '   x ->
    ' y 0,0--D--1,0
    ' | |  \    / |
    ' V A  centre C
    '   | /     \ |
    '   0,1--B--1,1

    ' get heights for left and right corners and the distances
    If onTileY > onTileX Then
        If onTileY < TILE_UNITS - onTileX Then
            ' A
            right_ = height(0, 0)
            left_ = height(0, 1)
            towardsCenter = onTileX
            towardsRight = TILE_UNITS - onTileY
        Else
            ' B
            right_ = height(0, 1)
            left_ = height(1, 1)
            towardsCenter = TILE_UNITS - onTileY
            towardsRight = TILE_UNITS - onTileX
        End If
    Else
        If onTileX > TILE_UNITS - onTileY Then
            ' C
            right_ = height(1, 1)
            left_ = height(1, 0)
            towardsCenter = TILE_UNITS - onTileX
            towardsRight = onTileY
        Else
            ' D
            right_ = height(1, 0)
            left_ = height(0, 0)
            towardsCenter = onTileY
            towardsRight = onTileX
        End If
    End If
    If Not (towardsCenter <= TILE_UNITS \ 2) Then ASSERT_FAILURE "towardsCenter <= (1<<7) / 2", "src/map.cpp:1696", "map_Height", "towardsCenter is too high"

    ' now we have:
    '    left   m    right
    '         center

    middle = (left_ + right_) \ 2
    onBottom = left_ * (TILE_UNITS - towardsRight) + right_ * towardsRight
    result = onBottom + (center - middle) * towardsCenter * 2

    map_Height = (result + TILE_UNITS \ 2) \ TILE_UNITS
End Function

' returns the max and min height of a tile by looking at the four corners
'   in tile coords
Public Sub getTileMaxMin(ByVal x As Long, ByVal y As Long, pMax As Long, pMin As Long)
    Dim i As Long, j As Long, height As Long
    pMin = &H7FFFFFFF
    pMax = &H80000000

    For j = 0 To 1
        For i = 0 To 1
            height = map_TileHeight(x + i, y + j)
            If height < pMin Then pMin = height
            If height > pMax Then pMax = height
        Next
    Next
End Sub

' Flood fill a "continent".
' TODO take into account scroll limits and update continents on scroll limit changes
' (VB) varContinent: 0 = MAPTILE::limitedContinent, 1 = MAPTILE::hoverContinent
Private Sub mapFloodFill(ByVal x As Long, ByVal y As Long, ByVal continent As Long, ByVal blockedBits As Long, ByVal varContinent As Long)
    Dim openX() As Long, openY() As Long, openCount As Long, posX As Long, posY As Long, i As Long
    Dim nposX As Long, nposY As Long, psTile As Long, c As Long

    ReDim openX(0 To 255)
    ReDim openY(0 To 255)
    openX(0) = x: openY(0) = y
    openCount = 1
    setContinent mapTile(x, y), varContinent, continent ' Set continent value

    Do While openCount > 0
        ' Pop the first open node off the list for this iteration
        openCount = openCount - 1
        posX = openX(openCount)
        posY = openY(openCount)

        ' Add accessible neighbouring tiles to the open list
        For i = 0 To NUM_DIR - 1
            ' rely on the fact that all border tiles are inaccessible to avoid checking explicitly
            nposX = posX + aDirOffset_x(i)
            nposY = posY + aDirOffset_y(i)

            If nposX < 1 Or nposY < 1 Or nposX > mapWidth - 2 Or nposY > mapHeight - 2 Then
                GoTo continue_
            End If
            psTile = mapTile(nposX, nposY)

            If varContinent = 0 Then c = psMapTiles(psTile).limitedContinent Else c = psMapTiles(psTile).hoverContinent
            If (blockTile(nposX, nposY, AUX_MAP) And blockedBits) = 0 And c = 0 Then
                If openCount > UBound(openX) Then
                    ReDim Preserve openX(0 To openCount * 2)
                    ReDim Preserve openY(0 To openCount * 2)
                End If
                openX(openCount) = nposX: openY(openCount) = nposY   ' add to open list
                openCount = openCount + 1
                setContinent psTile, varContinent, continent        ' Set continent value
            End If
continue_:
        Next
    Loop
End Sub

' (VB) psTile->*varContinent = continent
Private Sub setContinent(ByVal psTile As Long, ByVal varContinent As Long, ByVal continent As Long)
    If varContinent = 0 Then
        psMapTiles(psTile).limitedContinent = continent And &HFFFF&
    Else
        psMapTiles(psTile).hoverContinent = continent And &HFFFF&
    End If
End Sub

Public Sub mapFloodFillContinents()
    Dim x As Long, y As Long, limitedContinents As Long, hoverContinents As Long, psTile As Long

    ' Clear continents
    For y = 0 To mapHeight - 1
        For x = 0 To mapWidth - 1
            psTile = mapTile(x, y)

            psMapTiles(psTile).limitedContinent = 0
            psMapTiles(psTile).hoverContinent = 0
        Next
    Next

    ' Iterate over the whole map, looking for unset continents
    For y = 1 To mapHeight - 3
        For x = 1 To mapWidth - 3
            psTile = mapTile(x, y)

            If psMapTiles(psTile).limitedContinent = 0 And Not fpathBlockingTile(x, y, PROPULSION_TYPE_WHEELED) Then
                limitedContinents = limitedContinents + 1
                mapFloodFill x, y, limitedContinents, WATER_BLOCKED Or FEATURE_BLOCKED, 0
            ElseIf psMapTiles(psTile).limitedContinent = 0 And Not fpathBlockingTile(x, y, PROPULSION_TYPE_PROPELLOR) Then
                limitedContinents = limitedContinents + 1
                mapFloodFill x, y, limitedContinents, LAND_BLOCKED Or FEATURE_BLOCKED, 0
            End If

            If psMapTiles(psTile).hoverContinent = 0 And Not fpathBlockingTile(x, y, PROPULSION_TYPE_HOVER) Then
                hoverContinents = hoverContinents + 1
                mapFloodFill x, y, hoverContinents, FEATURE_BLOCKED, 1
            End If
        Next
    Next
    debug_ 1977, LOG_MAP, "mapFloodFillContinents", "Found %d limited and %d hover continents", limitedContinents, hoverContinents
End Sub

Public Sub tileSetFire(ByVal x As Long, ByVal y As Long, ByVal duration As Long)
    Dim posX As Long, posY As Long, tile As Long, currentTime As Long, fireEndTime As Long
    posX = map_coord(x)
    posY = map_coord(y)
    tile = mapTile(posX, posY)

    currentTime = Int(U32(gameTime) / GAME_TICKS_PER_UPDATE) And &HFFFF&
    fireEndTime = Int(U32(FromU32(U32(gameTime) + U32(duration))) / GAME_TICKS_PER_UPDATE) And &HFFFF&
    If currentTime = fireEndTime Then
        Exit Sub  ' Fire already ended.
    End If
    If (psMapTiles(tile).tileInfoBits And BITS_ON_FIRE) <> 0 And ((fireEndTime - currentTime) And &HFFFF&) < ((psMapTiles(tile).fireEndTime - currentTime) And &HFFFF&) Then
        Exit Sub  ' Tile already on fire, and that fire lasts longer.
    End If

    ' Burn, tile, burn!
    psMapTiles(tile).tileInfoBits = psMapTiles(tile).tileInfoBits Or BITS_ON_FIRE
    psMapTiles(tile).fireEndTime = fireEndTime

    syncDebug "tileSetFire", "Fire tile{%d, %d} dur%u end%d", posX, posY, duration, fireEndTime
End Sub

' Check if tile contained within the given world coordinates is burning.
Public Function fireOnLocation(ByVal x As Long, ByVal y As Long) As Boolean
    Dim posX As Long, posY As Long, psTile As Long
    posX = map_coord(x)
    posY = map_coord(y)
    psTile = mapTile(posX, posY)

    ' (ASSERT(psTile, ...): a tile index is never nullptr)
    fireOnLocation = TileIsBurning(psTile)
End Function

' ---- the danger maps

' This function runs in a separate thread! (VB: where the thread would run it, see the module's comment)
Private Function dangerFloodFill(ByVal player As Long) As Long
    Dim i As Long, pos As Vector2i, nposX As Long, nposY As Long, aux As Long, block As Long, x As Long, y As Long
    Dim start As Boolean
    pos = getPlayerStartPosition(player)
    start = True    ' hack to disregard the blocking status of any building exactly on the starting position

    ' Set our danger bits
    For y = 0 To mapHeight - 1
        For x = 0 To mapWidth - 1
            auxSet x, y, MAX_PLAYERS + AUX_DANGERMAP, AUXBITS_DANGER
            auxClear x, y, MAX_PLAYERS + AUX_DANGERMAP, AUXBITS_TEMPORARY
        Next
    Next

    pos.x = map_coord(pos.x)
    pos.y = map_coord(pos.y)
    bucketcounter = 0

    Do
        ' Add accessible neighbouring tiles to the open list
        For i = 0 To NUM_DIR - 1
            nposX = pos.x + aDirOffset_x(i)
            nposY = pos.y + aDirOffset_y(i)
            If Not tileOnMap(nposX, nposY) Then
                GoTo next_dir
            End If
            aux = auxTile(nposX, nposY, MAX_PLAYERS + AUX_DANGERMAP)
            block = blockTile(pos.x, pos.y, AUX_DANGERMAP)
            If (aux And AUXBITS_TEMPORARY) = 0 And (aux And AUXBITS_THREAT) = 0 And (aux And AUXBITS_DANGER) <> 0 Then
                ' Note that we do not consider water to be a blocker here. This may or may not be a feature...
                If (block And FEATURE_BLOCKED) = 0 And ((aux And AUXBITS_NONPASSABLE) = 0 Or start) Then
                    floodbucket_x(bucketcounter) = nposX And &HFF&
                    floodbucket_y(bucketcounter) = nposY And &HFF&
                    bucketcounter = bucketcounter + 1
                    If start And (aux And AUXBITS_NONPASSABLE) = 0 Then
                        start = False
                    End If
                Else
                    auxClear nposX, nposY, MAX_PLAYERS + AUX_DANGERMAP, AUXBITS_DANGER
                End If
                auxSet nposX, nposY, MAX_PLAYERS + AUX_DANGERMAP, AUXBITS_TEMPORARY ' make sure we do not process it more than once
            End If
next_dir:
        Next

        ' Clear danger
        auxClear pos.x, pos.y, MAX_PLAYERS + AUX_DANGERMAP, AUXBITS_DANGER

        ' Pop the last open node off the bucket list for the next iteration
        If bucketcounter <> 0 Then
            bucketcounter = bucketcounter - 1
            pos.x = floodbucket_x(bucketcounter)
            pos.y = floodbucket_y(bucketcounter)
        End If
    Loop While bucketcounter <> 0
    dangerFloodFill = 0
End Function

Private Sub threatUpdateTarget(ByVal player As Long, ByVal psObj As Long, ByVal ground_ As Boolean, ByVal air_ As Boolean)
    Dim i As Long, n As Long, pos As TILEPOS
    If BASE_OBJECT_visible(psObj, player) <> 0 Or SIMPLE_OBJECT_born(psObj) = 2 Then
        If psObj \ OBJ_HANDLE_UNIT = OBJ_DROID Then n = aDroid(psObj).watchedTiles_size Else n = aStruct(psObj).watchedTiles_size
        For i = 0 To n - 1
            If psObj \ OBJ_HANDLE_UNIT = OBJ_DROID Then pos = aDroid(psObj).watchedTiles(i) Else pos = aStruct(psObj).watchedTiles(i)
            If ground_ Then
                auxSet pos.x, pos.y, MAX_PLAYERS + AUX_DANGERMAP, AUXBITS_THREAT  ' set ground_ threat for this tile
            End If
            If air_ Then
                auxSet pos.x, pos.y, MAX_PLAYERS + AUX_DANGERMAP, AUXBITS_AATHREAT  ' set air_ threat for this tile
            End If
        Next
    End If
End Sub

Private Sub threatUpdate(ByVal player As Long)
    Dim i As Long, weapon_ As Long, x As Long, y As Long, mode As Long, it As Long, psDroid As Long, psStruct As Long

    ' Step 1: Clear our threat bits
    For y = 0 To mapHeight - 1
        For x = 0 To mapWidth - 1
            auxClear x, y, MAX_PLAYERS + AUX_DANGERMAP, AUXBITS_THREAT Or AUXBITS_AATHREAT
        Next
    Next

    ' Step 2: Set threat bits
    For i = 0 To MAX_PLAYERS - 1
        If aiCheckAlliances(player, i) Then
            ' No need to iterate friendly objects
            GoTo next_player
        End If

        it = apsDroidLists(i).head
        Do While it <> 0
            psDroid = lnObj(it)
            mode = 0

            Select Case aDroid(psDroid).droidType
            Case DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT, DROID_REPAIR, DROID_CYBORG_REPAIR
                GoTo next_droid  ' hack that really should not be needed, but is -- trucks can SHOOT_ON_GROUND...!
            End Select
            For weapon_ = 0 To aDroid(psDroid).numWeaps - 1
                mode = mode Or asWeaponStats(DROID_getWeaponStats(psDroid, weapon_)).surfaceToAir
            Next
            If aDroid(psDroid).droidType = DROID_SENSOR Then  ' special treatment for sensor turrets, no multiweapon support
                mode = mode Or SHOOT_ON_GROUND  ' assume it only shoots at ground targets for now
            End If
            If mode > 0 Then
                threatUpdateTarget player, psDroid, (mode And SHOOT_ON_GROUND) <> 0, (mode And SHOOT_IN_AIR) <> 0
            End If
next_droid:
            it = lnNext(it)
        Loop

        it = apsStructLists(i).head
        Do While it <> 0
            psStruct = lnObj(it)
            mode = 0

            For weapon_ = 0 To aStruct(psStruct).numWeaps - 1
                mode = mode Or asWeaponStats(BASE_OBJECT_getWeaponStats(psStruct, weapon_)).surfaceToAir
            Next
            If asStructureStats(aStruct(psStruct).pStructureType).pSensor <> -1 Then
                If asSensorStats(asStructureStats(aStruct(psStruct).pStructureType).pSensor).location = LOC_TURRET Then  ' special treatment for sensor turrets
                    mode = mode Or SHOOT_ON_GROUND  ' assume it only shoots at ground targets for now
                End If
            End If
            If mode > 0 Then
                threatUpdateTarget player, psStruct, (mode And SHOOT_ON_GROUND) <> 0, (mode And SHOOT_IN_AIR) <> 0
            End If
            it = lnNext(it)
        Loop
next_player:
    Next
End Sub

Public Sub mapInit()
    Dim player As Long

    Erase floodbucket_x
    Erase floodbucket_y
    If mapWidth * mapHeight > 0 Then
        ReDim floodbucket_x(0 To mapWidth * mapHeight - 1)
        ReDim floodbucket_y(0 To mapWidth * mapHeight - 1)
    End If

    lastDangerUpdate = 0
    lastDangerPlayer = -1

    ' Start danger thread (not used for campaign for now - mission map swaps too icky)
    ' (ASSERT(dangerSemaphore == nullptr && dangerThread == nullptr, ...): there is no thread)
    If game.typ = LEVEL_TYPE_SKIRMISH Then
        For player = 0 To MAX_PLAYERS - 1
            auxMapStore player, AUX_DANGERMAP
            threatUpdate player
            dangerFloodFill player
            auxMapRestore player, AUX_DANGERMAP, AUXBITS_DANGER Or AUXBITS_THREAT Or AUXBITS_AATHREAT
        Next
        lastDangerPlayer = 0
        ' (the thread starts: its first work)
        dangerFloodFill lastDangerPlayer
    End If
End Sub

Public Sub mapUpdate()
    Dim currentTime As Long, posX As Long, posY As Long, tile As Long

    currentTime = Int(U32(gameTime) / GAME_TICKS_PER_UPDATE) And &HFFFF&

    For posY = 0 To mapHeight - 1
        For posX = 0 To mapWidth - 1
            tile = mapTile(posX, posY)

            If (psMapTiles(tile).tileInfoBits And BITS_ON_FIRE) <> 0 And psMapTiles(tile).fireEndTime = currentTime Then
                ' Extinguish, tile, extinguish!
                psMapTiles(tile).tileInfoBits = psMapTiles(tile).tileInfoBits And Not BITS_ON_FIRE

                syncDebug "mapUpdate", "Extinguished tile{%d, %d}", posX, posY
            End If
        Next
    Next

    If ULT(FromU32(U32(lastDangerUpdate) + GAME_TICKS_FOR_DANGER), gameTime) And game.typ = LEVEL_TYPE_SKIRMISH Then
        syncDebug "mapUpdate", "Do danger maps."
        lastDangerUpdate = gameTime

        ' Lock if previous job not done yet
        ' (wzSemaphoreWait(dangerDoneSemaphore): done)

        auxMapRestore lastDangerPlayer, AUX_DANGERMAP, AUXBITS_THREAT Or AUXBITS_AATHREAT Or AUXBITS_DANGER
        lastDangerPlayer = (lastDangerPlayer + 1) Mod game.maxPlayers
        auxMapStore lastDangerPlayer, AUX_DANGERMAP
        threatUpdate lastDangerPlayer
        ' (wzSemaphorePost(dangerSemaphore): the thread's work)
        dangerFloodFill lastDangerPlayer
    End If
End Sub
