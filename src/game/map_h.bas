Attribute VB_Name = "map_h"
' src/map.h: the map's tile structure, flags and inline functions (the map
' itself is map_c's). The struct MAPTILE is MAPTILE_ (mapTile keeps its
' name); a MAPTILE * is the tile's index in psMapTiles (mapTile returns it). The aux maps (std::unique_ptr<uint8_t[]> arrays) are
' two-dimensional Byte arrays: psBlockMap(slot, tile), psAuxMap(player,
' tile). Not ported yet (they need the display or more of the objects:
' clusters B and E): SET_TILE_VISIBLE, hasSensorOnTile and the Vector2i
' overloads.
Option Explicit

Public Const TALLOBJECT_YMAX As Long = 200
Public Const TALLOBJECT_ADJUST As Long = 300

Public Const BITS_MARKED As Long = &H1          ' Is this tile marked?
Public Const BITS_DECAL As Long = &H2           ' Does this tile has a decal? If so, the tile from "texture" is drawn on top of the terrain.
Public Const BITS_FPATHBLOCK As Long = &H10     ' Bit set temporarily by find path to mark a blocking tile
Public Const BITS_ON_FIRE As Long = &H20        ' Whether tile is burning
Public Const BITS_GATEWAY As Long = &H40        ' Bit set to show a gateway on the tile

Public Type GROUND_TYPE
    textureName As String
    textureSize As Single
    normalMapTextureName As String
    specularMapTextureName As String
    heightMapTextureName As String
    highQualityTextures As Boolean      ' whether this ground_type has normal / specular / height maps
End Type

' Information stored with each tile
Public Type MAPTILE_
    tileInfoBits As Long                ' uint8_t
    tileExploredBits As Long            ' PlayerMask (uint16_t)
    sensorBits As Long                  ' bit per player, who can see tile with sensor
    watchers(0 To MAX_PLAYERS - 1) As Long  ' uint16_t: player sees through fog of war here with this many objects
    texture As Long                     ' uint16_t: Which graphics texture is on this tile
    height As Long                      ' The height at the top left of the tile
    psObject As Long                    ' BASE_OBJECT *: Any object sitting on the location (e.g. building)
    limitedContinent As Long            ' uint16_t: For land or sea limited propulsion types
    hoverContinent As Long              ' uint16_t: For hover type propulsions
    fireEndTime As Long                 ' uint16_t: The (uint16_t)(gameTime / GAME_TICKS_PER_UPDATE) that BITS_ON_FIRE should be cleared.
    waterLevel As Long                  ' At what height is the water for this tile
    jammerBits As Long                  ' bit per player, who is jamming tile
    sensors(0 To MAX_PLAYERS - 1) As Long   ' uint16_t: player sees this tile with this many radar sensors
    jammers(0 To MAX_PLAYERS - 1) As Long   ' uint16_t: player jams the tile with this many objects

    ' DISPLAY ONLY (NOT for use in game calculations)
    ground As Long                      ' uint8_t: The ground type used for the terrain renderer
    illumination As Long                ' uint8_t: How bright is this tile? = diffuseSunLight * ambientOcclusion
    ambientOcclusion As Long            ' uint8_t: ambient occlusion. from 1 (max occlusion) to 254 (no occlusion), similar to illumination.
    level As Single                     ' The visibility level of the top left of the tile, for this client. for terrain lightmap
End Type

Public Const MAX_GROUND_TYPES As Long = 12

Public Const AIR_BLOCKED As Long = &H1          ' Aircraft cannot pass tile
Public Const FEATURE_BLOCKED As Long = &H2      ' Ground units cannot pass tile due to item in the way
Public Const WATER_BLOCKED As Long = &H4        ' Units that cannot pass water are blocked by this tile
Public Const LAND_BLOCKED As Long = &H8         ' The inverse of the above -- for propeller driven crafts

Public Const AUXBITS_NONPASSABLE As Long = &H1  ' Is there any building blocking here, other than a gate that would open for us?
Public Const AUXBITS_OUR_BUILDING As Long = &H2 ' Do we or our allies have a building at this tile
Public Const AUXBITS_BLOCKING As Long = &H4     ' Is there any building currently blocking here?
Public Const AUXBITS_TEMPORARY As Long = &H8    ' Temporary bit used in calculations
Public Const AUXBITS_DANGER As Long = &H10      ' Does AI sense danger going there?
Public Const AUXBITS_THREAT As Long = &H20      ' Can hostile players shoot here?
Public Const AUXBITS_AATHREAT As Long = &H40    ' Can hostile players shoot at my VTOLs here?
Public Const AUXBITS_UNUSED As Long = &H80      ' Unused
Public Const AUXBITS_ALL As Long = &HFF

Public Const AUX_MAP As Long = 0
Public Const AUX_ASTARMAP As Long = 1
Public Const AUX_DANGERMAP As Long = 2
Public Const AUX_MAX As Long = 3

' Arbitrary maximum number of terrain textures - used in look up table for terrain type
Public Const MAX_TILE_TEXTURES As Long = 255

' Find aux bitfield for a given tile
Public Function auxTile(ByVal x As Long, ByVal y As Long, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS + AUX_MAX) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS + AUX_MAX", "src/map.h:121", "auxTile", "invalid player: %d", player: auxTile = AUXBITS_ALL: Exit Function
    auxTile = psAuxMap(player, x + y * mapWidth)
End Function

' Find blocking bitfield for a given tile
Public Function blockTile(ByVal x As Long, ByVal y As Long, ByVal slot As Long) As Long
    blockTile = psBlockMap(slot, x + y * mapWidth)
End Function

' Store a shadow copy of a player's aux map for use in threaded calculations
Public Sub auxMapStore(ByVal player As Long, ByVal slot As Long)
    Dim i As Long
    For i = 0 To mapWidth * mapHeight - 1
        psBlockMap(slot, i) = psBlockMap(0, i)
        psAuxMap(MAX_PLAYERS + slot, i) = psAuxMap(player, i)
    Next
End Sub

' Restore selected fields from the shadow copy of a player's aux map (ignoring the block map)
Public Sub auxMapRestore(ByVal player As Long, ByVal slot As Long, ByVal mask As Long)
    Dim i As Long, original As Long, cached As Long

    For i = 0 To mapHeight * mapWidth - 1
        original = psAuxMap(player, i)
        cached = psAuxMap(MAX_PLAYERS + slot, i)
        psAuxMap(player, i) = (original Xor ((original Xor cached) And mask)) And &HFF&
    Next
End Sub

' Set aux bits. Always set identically for all players. States not set are retained.
Public Sub auxSet(ByVal x As Long, ByVal y As Long, ByVal player As Long, ByVal state As Long)
    psAuxMap(player, x + y * mapWidth) = psAuxMap(player, x + y * mapWidth) Or (state And &HFF&)
End Sub

' Set aux bits. Always set identically for all players. States not set are retained.
Public Sub auxSetAll(ByVal x As Long, ByVal y As Long, ByVal state As Long)
    Dim i As Long

    For i = 0 To MAX_PLAYERS - 1
        psAuxMap(i, x + y * mapWidth) = psAuxMap(i, x + y * mapWidth) Or (state And &HFF&)
    Next
End Sub

' Clear aux bits. Always set identically for all players. States not cleared are retained.
Public Sub auxClear(ByVal x As Long, ByVal y As Long, ByVal player As Long, ByVal state As Long)
    psAuxMap(player, x + y * mapWidth) = psAuxMap(player, x + y * mapWidth) And (Not state) And &HFF&
End Sub

' Clear all aux bits. Always set identically for all players. States not cleared are retained.
Public Sub auxClearAll(ByVal x As Long, ByVal y As Long, ByVal state As Long)
    Dim i As Long

    For i = 0 To MAX_PLAYERS - 1
        psAuxMap(i, x + y * mapWidth) = psAuxMap(i, x + y * mapWidth) And (Not state) And &HFF&
    Next
End Sub

' Set aux bits. Always set identically for all players. States not set are retained.
Public Sub auxSetAllied(ByVal x As Long, ByVal y As Long, ByVal player As Long, ByVal state As Long)
    Dim i As Long

    For i = 0 To MAX_PLAYERS - 1
        If aiCheckAlliances(player, i) Then
            psAuxMap(i, x + y * mapWidth) = psAuxMap(i, x + y * mapWidth) Or (state And &HFF&)
        End If
    Next
End Sub

' Set aux bits. Always set identically for all players. States not set are retained.
Public Sub auxSetEnemy(ByVal x As Long, ByVal y As Long, ByVal player As Long, ByVal state As Long)
    Dim i As Long

    For i = 0 To MAX_PLAYERS - 1
        If Not aiCheckAlliances(player, i) Then
            psAuxMap(i, x + y * mapWidth) = psAuxMap(i, x + y * mapWidth) Or (state And &HFF&)
        End If
    Next
End Sub

' Set blocking bits. Always set identically for all players. States not set are retained.
Public Sub auxSetBlocking(ByVal x As Long, ByVal y As Long, ByVal state As Long)
    psBlockMap(0, x + y * mapWidth) = psBlockMap(0, x + y * mapWidth) Or (state And &HFF&)
End Sub

' Clear blocking bits. Always set identically for all players. States not cleared are retained.
Public Sub auxClearBlocking(ByVal x As Long, ByVal y As Long, ByVal state As Long)
    psBlockMap(0, x + y * mapWidth) = psBlockMap(0, x + y * mapWidth) And (Not state) And &HFF&
End Sub

' Check if tile contains a structure or feature. Function is thread-safe,
' but do not rely on the result if you mean to alter the object pointer.
Public Function TileIsOccupied(ByVal tile As Long) As Boolean
    TileIsOccupied = psMapTiles(tile).psObject <> 0
End Function

Public Function TileIsKnownOccupied(ByVal tile As Long, ByVal player As Long) As Boolean
    Dim psObj As Long
    If Not TileIsOccupied(tile) Then Exit Function
    psObj = psMapTiles(tile).psObject
    If SIMPLE_OBJECT_type(psObj) <> OBJ_STRUCTURE Then
        TileIsKnownOccupied = True
    ElseIf aStruct(psObj).visible(player) <> 0 Then
        TileIsKnownOccupied = True
    Else
        TileIsKnownOccupied = aiCheckAlliances(player, aStruct(psObj).player)
    End If
End Function

' Check if tile contains a structure. Function is NOT thread-safe.
Public Function TileHasStructure(ByVal tile As Long) As Boolean
    If TileIsOccupied(tile) Then TileHasStructure = SIMPLE_OBJECT_type(psMapTiles(tile).psObject) = OBJ_STRUCTURE
End Function

' Check if tile contains a feature. Function is NOT thread-safe.
Public Function TileHasFeature(ByVal tile As Long) As Boolean
    If TileIsOccupied(tile) Then TileHasFeature = SIMPLE_OBJECT_type(psMapTiles(tile).psObject) = OBJ_FEATURE
End Function

' Check if tile contains a wall structure. Function is NOT thread-safe.
Public Function TileHasWall(ByVal tile As Long) As Boolean
    If TileHasStructure(tile) Then
        Select Case asStructureStats(aStruct(psMapTiles(tile).psObject).pStructureType).typ
        Case REF_WALL, REF_GATE, REF_WALLCORNER
            TileHasWall = True
        End Select
    End If
End Function

' Check if tile contains a wall structure. Function is NOT thread-safe.
' This function is specifically for raycast callback, because "TileHasWall" is used
' to decide wether or not it's possible to place a new defense structure on top of a wall.
' We do not want to allow place more defense structures on top of already existing ones.
Public Function TileHasWall_raycast(ByVal tile As Long) As Boolean
    Dim psStats As Long
    If TileHasStructure(tile) Then
        psStats = aStruct(psMapTiles(tile).psObject).pStructureType
        Select Case asStructureStats(psStats).typ
        Case REF_WALL, REF_GATE, REF_WALLCORNER
            TileHasWall_raycast = True
        Case Else
            ' hadrcrete towers are technically a WALL + weapon, so count them as walls too
            TileHasWall_raycast = asStructureStats(psStats).combinesWithWall
        End Select
    End If
End Function

' Check if tile contains a small structure. Function is NOT thread-safe.
Public Function TileHasSmallStructure(ByVal tile As Long) As Boolean
    If TileHasStructure(tile) Then TileHasSmallStructure = asStructureStats(aStruct(psMapTiles(tile).psObject).pStructureType).height = 1
End Function

' psDroid->visible is for objects, this is for tiles (DISPLAY ONLY)
Public Function tileIsClearlyVisible(ByVal psTile As Long) As Boolean
    If Not ULT(selectedPlayer, MAX_PLAYERS) Or godMode Then
        tileIsClearlyVisible = True
        Exit Function
    End If
    tileIsClearlyVisible = (psMapTiles(psTile).sensorBits And BIT_(selectedPlayer)) <> 0
End Function

' Check if tile is burning.
Public Function TileIsBurning(ByVal tile As Long) As Boolean
    TileIsBurning = (psMapTiles(tile).tileInfoBits And BITS_ON_FIRE) <> 0
End Function

' Check if tile has been explored.
Public Function tileIsExplored(ByVal psTile As Long) As Boolean
    If ULT(selectedPlayer, MAX_PLAYERS) = False Then tileIsExplored = True: Exit Function
    tileIsExplored = (psMapTiles(psTile).tileExploredBits And BIT_(selectedPlayer)) <> 0
End Function

Public Sub SET_TILE_DECAL(ByVal x As Long)
    psMapTiles(x).tileInfoBits = psMapTiles(x).tileInfoBits Or BITS_DECAL
End Sub

Public Sub CLEAR_TILE_DECAL(ByVal x As Long)
    psMapTiles(x).tileInfoBits = psMapTiles(x).tileInfoBits And (Not BITS_DECAL) And &HFF&
End Sub

Public Function TILE_HAS_DECAL(ByVal x As Long) As Boolean
    TILE_HAS_DECAL = (psMapTiles(x).tileInfoBits And BITS_DECAL) <> 0
End Function

' Allows us to do if(TRI_FLIPPED(psTile))
Public Function TRI_FLIPPED(ByVal x As Long) As Boolean
    TRI_FLIPPED = (psMapTiles(x).texture And TILE_TRIFLIP) <> 0
End Function

' Flips the triangle partition on a tile pointer
Public Sub TOGGLE_TRIFLIP(ByVal x As Long)
    psMapTiles(x).texture = psMapTiles(x).texture Xor TILE_TRIFLIP
End Sub

' Can player number p has explored tile t?
Public Function TEST_TILE_VISIBLE(ByVal p As Long, ByVal t As Long) As Boolean
    TEST_TILE_VISIBLE = (psMapTiles(t).tileExploredBits And BIT_(p)) <> 0
End Function

' Can selectedPlayer see tile t?
' To be used for *DISPLAY* purposes only (*not* game-state/calculation related)
Public Function TEST_TILE_VISIBLE_TO_SELECTEDPLAYER(ByVal pTile As Long) As Boolean
    If godMode Then
        ' always visible
        TEST_TILE_VISIBLE_TO_SELECTEDPLAYER = True
        Exit Function
    End If
    If Not ULT(selectedPlayer, MAX_PLAYERS) Then
        TEST_TILE_VISIBLE_TO_SELECTEDPLAYER = False
        Exit Function
    End If
    TEST_TILE_VISIBLE_TO_SELECTEDPLAYER = TEST_TILE_VISIBLE(selectedPlayer, pTile)
End Function

Public Function terrainType(ByVal tile As Long) As Long
    terrainType = terrainTypes(TileNumber_tile(psMapTiles(tile).texture))
End Function

' Make sure world coordinates are inside the map
' Clip world coordinates to make sure they're inside the map's boundaries
'  \param worldX a pointer to a X coordinate inside the map
'  \param worldY a pointer to a Y coordinate inside the map
'  \post 1 <= *worldX <= world_coord(mapWidth)-1 and
'        1 <= *worldy <= world_coord(mapHeight)-1
Public Sub clip_world_offmap(worldX As Long, worldY As Long)
    ' x,y must be > 0
    If worldX < 1 Then worldX = 1
    If worldY < 1 Then worldY = 1
    If world_coord(mapWidth) - 1 < worldX Then worldX = world_coord(mapWidth) - 1
    If world_coord(mapHeight) - 1 < worldY Then worldY = world_coord(mapHeight) - 1
End Sub

' Return a pointer to the tile structure at x,y in map coordinates
Public Function mapTile(ByVal x As Long, ByVal y As Long) As Long
    ' Clamp x and y values to actual ones
    ' Give one tile worth of leeway before asserting, for units/transporters coming in from off-map.
    If Not (x >= -1) Then ASSERT_FAILURE "x >= -1", "src/map.h:438", "mapTile", "mapTile: x value is too small (%d,%d) in %dx%d", x, y, mapWidth, mapHeight
    If Not (y >= -1) Then ASSERT_FAILURE "y >= -1", "src/map.h:439", "mapTile", "mapTile: y value is too small (%d,%d) in %dx%d", x, y, mapWidth, mapHeight
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    If Not (x < mapWidth + 1) Then ASSERT_FAILURE "x < mapWidth + 1", "src/map.h:442", "mapTile", "mapTile: x value is too big (%d,%d) in %dx%d", x, y, mapWidth, mapHeight
    If Not (y < mapHeight + 1) Then ASSERT_FAILURE "y < mapHeight + 1", "src/map.h:443", "mapTile", "mapTile: y value is too big (%d,%d) in %dx%d", x, y, mapWidth, mapHeight
    If x > mapWidth - 1 Then x = mapWidth - 1
    If y > mapHeight - 1 Then y = mapHeight - 1

    mapTile = x + (y * mapWidth)
End Function

' Return a pointer to the tile structure at x,y in world coordinates
Public Function worldTile(ByVal x As Long, ByVal y As Long) As Long
    worldTile = mapTile(map_coord(x), map_coord(y))
End Function

' Return ground height of top-left corner of tile at x,y
Public Function map_TileHeight(ByVal x As Long, ByVal y As Long) As Long
    If x >= mapWidth Or y >= mapHeight Or x < 0 Or y < 0 Then
        map_TileHeight = 0
        Exit Function
    End If
    map_TileHeight = psMapTiles(x + (y * mapWidth)).height
End Function

' Return water height of top-left corner of tile at x,y
Public Function map_WaterHeight(ByVal x As Long, ByVal y As Long) As Long
    If x >= mapWidth Or y >= mapHeight Or x < 0 Or y < 0 Then
        map_WaterHeight = 0
        Exit Function
    End If
    map_WaterHeight = psMapTiles(x + (y * mapWidth)).waterLevel
End Function

' Return max(ground, water) height of top-left corner of tile at x,y
Public Function map_TileHeightSurface(ByVal x As Long, ByVal y As Long) As Long
    Dim i As Long
    If x >= mapWidth Or y >= mapHeight Or x < 0 Or y < 0 Then
        map_TileHeightSurface = 0
        Exit Function
    End If
    i = x + (y * mapWidth)
    If psMapTiles(i).height > psMapTiles(i).waterLevel Then
        map_TileHeightSurface = psMapTiles(i).height
    Else
        map_TileHeightSurface = psMapTiles(i).waterLevel
    End If
End Function


' sets the tile height
Public Sub setTileHeight(ByVal x As Long, ByVal y As Long, ByVal height As Long)
    If Not (x < mapWidth And x >= 0) Then ASSERT_FAILURE "x < mapWidth && x >= 0", "src/map.h:499", "setTileHeight", "x coordinate %d bigger than map width %u", x, mapWidth: Exit Sub
    If Not (y < mapHeight And x >= 0) Then ASSERT_FAILURE "y < mapHeight && x >= 0", "src/map.h:500", "setTileHeight", "y coordinate %d bigger than map height %u", y, mapHeight: Exit Sub

    psMapTiles(x + (y * mapWidth)).height = height
    markTileDirty x, y
End Sub

' Return whether a tile coordinate is on the map
Public Function tileOnMap(ByVal x As Long, ByVal y As Long) As Boolean
    tileOnMap = (x >= 0) And (x < mapWidth) And (y >= 0) And (y < mapHeight)
End Function

' Return whether a world coordinate is on the map
Public Function worldOnMap(ByVal x As Long, ByVal y As Long) As Boolean
    worldOnMap = (x >= 0) And (x < world_coord(mapWidth)) And _
                 (y >= 0) And (y < world_coord(mapHeight))
End Function

Public Sub makeTileRubbleTexture(ByVal psTile As Long, ByVal x As Long, ByVal y As Long, ByVal newTexture As Long)
    psMapTiles(psTile).texture = (TileNumber_texture(psMapTiles(psTile).texture) Or newTexture) And &HFFFF&
    SET_TILE_DECAL psTile
    markTileDirty x, y
End Sub
