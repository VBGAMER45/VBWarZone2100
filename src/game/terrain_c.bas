Attribute VB_Name = "terrain_c"
' src/terrain.cpp, terrain.h, terrain_defs.h: draws the terrain.
' It uses the ground property of every MAPTILE to divide the terrain into different layers.
' For every layer the GROUND_TYPE (from psGroundTypes) determines the texture and size.
' The technique used it called "texture splatting".
' Every layer only draws the spots where that terrain is, and additive blending is used to make transitions smooth.
' Decals are a kind of hack now, as for some tiles (where decal == true) the old tile is just drawn.
' The water is drawn using the hardcoded page-80 and page-81 textures.
' (VB) The VBOs are WzGfx's gfx_api_buffer objects, filled from VB arrays by
' address; the PSO templates take the pso_type; a TerrainShaderQuality is a
' Long (the enum's names would clash with the variable terrainShaderQuality).
' ShadowCascadesInfo is src/shared/gfx's shadows_h. The light map's pixels are kept in
' lightmapBytes, a copy of lightmapPixmap's buffer that updateLightMap writes and
' copies into the image at once (C++ writes the image's bytes one by one).
' With no graphics context at all (tools/run_tests.py's --vbtest runs, whose
' reference stubs loadTerrainTextures) there is nothing to load.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' terrain_defs.h: enum TerrainShaderQuality
Public Const UNINITIALIZED_PICK_DEFAULT As Long = -1 ' a non-selectable option for first-run - pick a default based on the current system settings
Public Const CLASSIC As Long = 0             ' classic, pixel-art, tile-based textures
Public Const MEDIUM As Long = 1              ' the mode used by at least WZ 3.2.x - 4.3.x
Public Const NORMAL_MAPPING As Long = 2      ' the highest-quality mode, which adds normal / specular / height maps and advanced lighting
Public Const TerrainShaderQuality_MAX As Long = NORMAL_MAPPING

' terrain.h
Public Enum TerrainShaderType
    FALLBACK                            ' old multi-pass method, which only supports "classic" rendering
    SINGLE_PASS                         ' new terrain rendering method, supports all TerrainShaderQuality modes
End Enum

' A sector contains all information to draw a square piece of the map.
' The actual geometry and texture data is not stored in here but in large VBO's.
' The sector only stores the index and length of the pieces it's going to use.
Private Type Sector
    geometryOffset As Long              ' The point in the geometry VBO where our geometry starts
    geometrySize As Long                ' The size of our geometry
    geometryIndexOffset As Long         ' The point in the index VBO where our triangles start
    geometryIndexSize As Long           ' The size of our indices
    waterOffset As Long                 ' The point in the water VBO where the water geometry starts
    waterSize As Long                   ' The size of the water geometry
    waterIndexOffset As Long            ' The point in the water index VBO where the water triangles start
    waterIndexSize As Long              ' The size of our water triangles

    decalOffset As Long                 ' = 0: Index into the decal VBO
    decalSize As Long                   ' = 0: Size of the part of the decal VBO we are going to use
    terrainAndDecalOffset As Long       ' = 0
    terrainAndDecalSize As Long         ' = 0
    draw As Boolean                     ' Do we draw this sector this frame?
    dirty As Boolean                    ' Do we need to update the geometry for this sector?
End Type

' VBO for gfx_api::TerrainLayer and TerrainDepth
Private Type TerrainVertex
    pos As Vector3f                     ' = Vector3f(0.f, 0.f, 0.f)
End Type

' A vertex with a position and texture coordinates
Private Type DecalVertex
    pos As Vector3f                     ' = Vector3f(0.f, 0.f, 0.f)
    uv As Vector2f                      ' = Vector2f(0.f, 0.f)
End Type

' using WaterVertex = glm::vec4; // w is depth: a Vector4f

' The lightmap texture
Private lightmap_texture As gfx_api_texture
' When are we going to update the lightmap next?
Private lightmapLastUpdate As Long      ' unsigned int
' How big is the lightmap?
Private lightmapWidth As Long           ' size_t
Private lightmapHeight As Long          ' size_t
' Lightmap image
Private lightmapPixmap As iV_Image
Private lightmapBytes() As Byte         ' (VB) lightmapPixmap's pixels, written by updateLightMap
' Lightmap values
Private Type LightmapCalculatedValues
    paramsXLight As Vector4f            ' = glm::vec4()
    paramsYLight As Vector4f            ' = glm::vec4()
    lightMatrix As mat4                 ' = glm::mat4(): shift the lightmap half a tile as lights are supposed to be placed at the center of a tile
    ModelUVLightmap As mat4             ' = glm::mat4()
End Type
Private lightmapValues As LightmapCalculatedValues
' Ticks per lightmap refresh
Private Const LIGHTMAP_REFRESH As Long = 80

' VBOs
Private geometryVBO As gfx_api_buffer, geometryIndexVBO As gfx_api_buffer
' VBOs
Private waterVBO As gfx_api_buffer, waterIndexVBO As gfx_api_buffer
Private terrainDecalVBO As gfx_api_buffer
' The amount we shift the water textures so the waves appear to be moving
Private waterOffset As Single

' These are properties of your videocard and hardware
Private GLmaxElementsVertices As Long, GLmaxElementsIndices As Long

' The sectors are stored here
Private sectors() As Sector
Private sectors_allocated As Boolean    ' (VB) sectors != nullptr
' The default sector size (a sector is sectorSize x sectorSize)
Private sectorSize As Long              ' = 15
' What is the distance we can see
Private terrainDistance As Long
' How many sectors have we actually got?
Private xSectors As Long, ySectors As Long

' Did we initialise the terrain renderer yet?
Private terrainInitialised As Boolean   ' = false

Private terrainDecalVertexUpdateBuffer() As TerrainDecalVertex
Private terrainDecalVertexUpdateBuffer_size As Long

' Helper variables for the DrawRangeElements functions
Private dreStart As Long, dreEnd As Long, dreOffset As Long ' GLuint
Private dreCount As Long                ' GLsizei
' Are we actually drawing something using the DrawRangeElements functions?
Private drawRangeElementsStarted As Boolean ' = false

Private terrainShaderQuality As Long    ' TerrainShaderQuality = UNINITIALIZED_PICK_DEFAULT
Private initializedTerrainShaderType As Boolean ' = false
Private maxTerrainMappingTextureSize As Long ' = 0: uninitialized - pick default on first run

Private Const MIN_TERRAIN_TEXTURE_SIZE As Long = 512

Private groundTexArr As gfx_api_texture_array
Private groundNormalArr As gfx_api_texture_array
Private groundSpecularArr As gfx_api_texture_array
Private groundHeightArr As gfx_api_texture_array

Private Type WaterTextures_Normal
    tex1 As gfx_api_texture             ' = nullptr
    tex2 As gfx_api_texture             ' = nullptr
End Type

Private Type WaterTextures_High
    tex As gfx_api_texture_array        ' = nullptr
    tex_nm As gfx_api_texture_array     ' = nullptr
    tex_sm As gfx_api_texture_array     ' = nullptr
End Type

Private waterTexturesNormal As WaterTextures_Normal
Private waterTexturesHigh As WaterTextures_High
Private waterClassicTexture As gfx_api_texture ' = nullptr: only used for classic mode

' (VB) the big uniform blocks, filled each frame (module variables, not 14 KB on the stack)
Private tcUniforms As TerrainCombinedUniforms
Private whUniforms As constant_buffer_type_SHADER_WATER_HIGH

' (VB) the C++ initialisers (Sub Main calls it)
Public Sub terrain_c_init()
    sectorSize = 15
    terrainShaderQuality = UNINITIALIZED_PICK_DEFAULT
    maxTerrainMappingTextureSize = 0
End Sub

' (VB) the render state's fields (getCurrentRenderState(): WzGfx gives them ByRef)
Private Sub renderState_get(fogEnabled As Boolean, fogBegin As Single, fogEnd As Single, fogColour As Long)
    Dim fog As Boolean, texPage As Long, rendMode As REND_MODE
    getCurrentRenderState fogEnabled, fog, fogColour, fogBegin, fogEnd, texPage, rendMode
End Sub

' Pass all remaining triangles to OpenGL
Private Sub finishDrawRangeElements(ByVal pso As pso_type)
    If drawRangeElementsStarted And dreCount > 0 Then
        If Not (dreEnd - dreStart + 1 <= GLmaxElementsVertices) Then ASSERT_FAILURE "dreEnd - dreStart + 1 <= GLmaxElementsVertices", "src/terrain.cpp:185", "finishDrawRangeElements", "too many vertices (%i)", dreEnd - dreStart + 1
        If Not (dreCount <= GLmaxElementsIndices) Then ASSERT_FAILURE "dreCount <= GLmaxElementsIndices", "src/terrain.cpp:186", "finishDrawRangeElements", "too many indices (%i)", dreCount
        gfx_api_pso_get(pso).draw_elements dreCount, 4 * dreOffset   ' sizeof(GLuint)*dreOffset
    End If
    drawRangeElementsStarted = False
End Sub

' Either draw the elements or batch them to be sent to OpenGL later
' This improves performance by reducing the amount of OpenGL calls.
Private Sub addDrawRangeElements(ByVal pso As pso_type, ByVal start As Long, ByVal end_ As Long, ByVal count As Long, ByVal offset As Long)
    If end_ - start + 1 > GLmaxElementsVertices Then
        debug_ 205, LOG_WARNING, "addDrawRangeElements", "A single call provided too much vertices, will operate at reduced performance or crash. Decrease the sector size to fix this."
    End If
    If count > GLmaxElementsIndices Then
        debug_ 209, LOG_WARNING, "addDrawRangeElements", "A single call provided too much indices, will operate at reduced performance or crash. Decrease the sector size to fix this."
    End If

    If Not drawRangeElementsStarted Then
        dreStart = start
        dreEnd = end_
        dreCount = count
        dreOffset = offset
        drawRangeElementsStarted = True
        Exit Sub
    End If

    ' check if we can append theoretically and
    ' check if this will not go over the bounds advised by the opengl implementation
    If dreOffset + dreCount <> offset Or _
       dreCount + count > GLmaxElementsIndices Or _
       end_ - dreStart + 1 > GLmaxElementsVertices Then
        finishDrawRangeElements pso
        ' start anew
        addDrawRangeElements pso, start, end_, count, offset
    Else
        ' OK to append
        dreCount = dreCount + count
        dreEnd = end_
    End If
    ' make sure we did everything right
    If Not (dreEnd - dreStart + 1 <= GLmaxElementsVertices) Then ASSERT_FAILURE "dreEnd - dreStart + 1 <= GLmaxElementsVertices", "src/terrain.cpp:239", "addDrawRangeElements", "too many vertices (%i)", dreEnd - dreStart + 1
    If Not (dreCount <= GLmaxElementsIndices) Then ASSERT_FAILURE "dreCount <= GLmaxElementsIndices", "src/terrain.cpp:240", "addDrawRangeElements", "too many indices (%i)", dreCount
End Sub

' (VB) std::swap of two Vector2f
Private Sub swapV2(a As Vector2f, b As Vector2f)
    Dim t As Vector2f
    t = a
    a = b
    b = t
End Sub

' NOTE:  The current (max) texture size of a tile is 128x128.  We allow up to a user defined texture size
' of 2048.  This will cause ugly seams for the decals, if user picks a texture size bigger than the tile!

Private Sub flipRotateTexCoords(ByVal texture As Long, sP1 As Vector2f, sP2 As Vector2f, sP3 As Vector2f, sP4 As Vector2f)
    Dim sPTemp As Vector2f
    If (texture And TILE_XFLIP) <> 0 Then
        swapV2 sP1, sP2
        swapV2 sP3, sP4
    End If
    If (texture And TILE_YFLIP) <> 0 Then
        swapV2 sP1, sP4
        swapV2 sP2, sP3
    End If

    Select Case (texture And TILE_ROTMASK) \ BIT_(TILE_ROTSHIFT)
    Case 1
        sPTemp = sP1
        sP1 = sP4
        sP4 = sP3
        sP3 = sP2
        sP2 = sPTemp
    Case 2
        sPTemp = sP1
        sP1 = sP3
        sP3 = sPTemp
        sPTemp = sP4
        sP4 = sP2
        sP2 = sPTemp
    Case 3
        sPTemp = sP1
        sP1 = sP2
        sP2 = sP3
        sP3 = sP4
        sP4 = sPTemp
    End Select
End Sub

' set up tile/decal coords for texture arrays. just flip/rotate, without borders and tex atlas stuff.
' (VB) uv: Vector2f uv[2][2] flattened, uv[i][j] is uv(i * 2 + j)
Private Function getTileTexArrCoords(uv() As Vector2f, ByVal tileNumber As Long) As Vector2f
    Dim texture As Long, sP(0 To 3) As Vector2f
    ' unmask proper values from compressed data
    texture = TileNumber_texture(tileNumber)
    sP(0) = glm_vec2(0!, 0!)
    sP(1) = glm_vec2(1!, 0!)
    sP(2) = glm_vec2(1!, 1!)
    sP(3) = glm_vec2(0!, 1!)
    flipRotateTexCoords texture, sP(0), sP(1), sP(2), sP(3)
    uv(0 + 0) = sP(0)
    uv(0 + 2) = sP(1)
    uv(1 + 2) = sP(2)
    uv(1 + 0) = sP(3)
    getTileTexArrCoords = glm_vec2(0.5!, 0.5!)
End Function

' Average the four positions to get the center
Private Sub averagePos(center As Vector3i, a As Vector3i, b As Vector3i, c As Vector3i, d As Vector3i)
    center.x = (a.x + b.x + c.x + d.x) \ 4
    center.y = (a.y + b.y + c.y + d.y) \ 4
    center.z = (a.z + b.z + c.z + d.z) \ 4
End Sub

' (VB) tileOnMap(x, y) && terrainType(mapTile(x, y)) == TER_WATER, short-circuited as C++ does
Private Function isWaterTile(ByVal x As Long, ByVal y As Long) As Boolean
    If tileOnMap(x, y) Then isWaterTile = (terrainType(mapTile(x, y)) = TER_WATER)
End Function

' Is this position next to a water tile?
Private Function isWater(ByVal x As Long, ByVal y As Long) As Boolean
    Dim result As Boolean
    result = False
    result = result Or isWaterTile(x, y)
    result = result Or isWaterTile(x - 1, y)
    result = result Or isWaterTile(x, y - 1)
    result = result Or isWaterTile(x - 1, y - 1)
    isWater = result
End Function

' Is this position an *actual* water-only tile?
Private Function isOnlyWater(ByVal x As Long, ByVal y As Long) As Boolean
    Dim result As Boolean
    result = True
    If result Then result = isWaterTile(x, y)
    If result Then result = isWaterTile(x - 1, y)
    If result Then result = isWaterTile(x, y - 1)
    If result Then result = isWaterTile(x - 1, y - 1)
    isOnlyWater = result
End Function

' Get the position of a grid point
Private Sub getGridPos(result As Vector3i, ByVal x As Long, ByVal y As Long, ByVal center As Boolean, ByVal water As Boolean)
    Dim a As Vector3i, b As Vector3i, c As Vector3i, d As Vector3i
    If center Then
        getGridPos a, x, y, False, water
        getGridPos b, x + 1, y, False, water
        getGridPos c, x, y + 1, False, water
        getGridPos d, x + 1, y + 1, False, water
        averagePos result, a, b, c, d
        Exit Sub
    End If
    result.x = world_coord(x)
    result.z = world_coord(-y)

    If x <= 0 Or y <= 0 Or x >= mapWidth Or y >= mapHeight Then
        result.y = 0
    Else
        If terrainShaderQuality <> CLASSIC Then
            result.y = map_TileHeight(x, y)
            If water Then
                result.y = map_WaterHeight(x, y)
            End If
        Else
            result.y = map_TileHeightSurface(x, y)
        End If
    End If
End Sub

Private Function getGridPosf(ByVal x As Long, ByVal y As Long, Optional ByVal center As Boolean = False, Optional ByVal water As Boolean = False) As Vector3f
    Dim r As Vector3i
    getGridPos r, x, y, center, water
    getGridPosf = glm_vec3(r.x, r.y, r.z)
End Function

' (VB) acos, as C's (an argument rounded past +-1 gives 0 or pi here, NaN there)
Private Function acos_(ByVal x As Double) As Double
    If IsNaNF(CSng(x)) Then
        acos_ = x                       ' (a NaN stays one, without an FPU comparison)
    ElseIf x >= 1# Then
        acos_ = 0#
    ElseIf x <= -1# Then
        acos_ = 3.14159265358979
    Else
        acos_ = Atn(-x / Sqr(-x * x + 1#)) + 2# * Atn(1#)
    End If
End Function

' (VB) getGridNormal's lambda calcNormal: p has 8 points
Private Function calcNormal(pc As Vector3f, p() As Vector3f) As Vector3f
    Dim res As Vector3f, i As Long, e1 As Vector3f, e2 As Vector3f, ang As Single, l1 As Single, l2 As Single
    res = glm_vec3(0!, 0!, 0!)
    For i = 0 To 7
        e1 = glm_vec3_sub(p((i + 1) Mod 8), pc)
        e2 = glm_vec3_sub(p(i), pc)
        l1 = glm_length3(e1)
        l2 = glm_length3(e2)
        ang = acos_(fdiv(fdiv(glm_dot3(e1, e2), l1), l2))
        res = glm_vec3_add(res, glm_vec3_mul(glm_cross(e1, e2), ang)) ' += normal * (2*area) * angle
    Next
    res = glm_normalize3(res)
    calcNormal = glm_vec3(-res.x, -res.y, -res.z)
End Function

' Get normal vector of grid point
Private Function getGridNormal(ByVal x As Long, ByVal y As Long, Optional ByVal center As Boolean = False) As Vector3f
    Dim p(0 To 7) As Vector3f, s As Vector3f, pc As Vector3f
    If center Then
        ' avg nearest normals provide better results
        s = glm_vec3_add(glm_vec3_add(glm_vec3_add(getGridNormal(x, y), getGridNormal(x + 1, y)), getGridNormal(x + 1, y + 1)), getGridNormal(x, y + 1))
        getGridNormal = glm_vec3_mul(s, 1! / 4!)
    Else
        pc = getGridPosf(x, y)
        p(0) = getGridPosf(x + 1, y)
        p(1) = getGridPosf(x, y, True)
        p(2) = getGridPosf(x, y + 1)
        p(3) = getGridPosf(x - 1, y, True)
        p(4) = getGridPosf(x - 1, y)
        p(5) = getGridPosf(x - 1, y - 1, True)
        p(6) = getGridPosf(x, y - 1)
        p(7) = getGridPosf(x, y - 1, True)
        getGridNormal = calcNormal(pc, p)
    End If
End Function

' Set the terrain and water geometry for the specified sector
Private Sub setSectorGeometry(ByVal sx As Long, ByVal sy As Long, _
                              geometry() As TerrainVertex, water() As Vector4f, _
                              geometrySize As Long, waterSize As Long)
    Dim x As Long, y As Long, pos As Vector3f, waterHeight As Single, wy As Single
    For x = sx * sectorSize To (sx + 1) * sectorSize
        For y = sy * sectorSize To (sy + 1) * sectorSize
            ' set up geometry
            pos = getGridPosf(x, y, False, False)
            geometry(geometrySize).pos = pos
            geometrySize = geometrySize + 1

            waterHeight = map_WaterHeight(x, y)
            If terrainShaderQuality <> CLASSIC Then wy = waterHeight Else wy = pos.y
            water(waterSize) = glm_vec4(pos.x, wy, pos.z, waterHeight - pos.y)
            waterSize = waterSize + 1

            pos = getGridPosf(x, y, True, False)
            geometry(geometrySize).pos = pos
            geometrySize = geometrySize + 1

            If terrainShaderQuality <> CLASSIC Then wy = waterHeight Else wy = pos.y
            water(waterSize) = glm_vec4(pos.x, wy, pos.z, waterHeight - pos.y)
            waterSize = waterSize + 1
        Next
    Next
End Sub

Private Sub setSectorDecalVertex_SinglePass(ByVal x As Long, ByVal y As Long, terrainDecalData() As TerrainDecalVertex, terrainDecalSize As Long)
    Dim pos As Vector3i, uv(0 To 3) As Vector2f, center As Vector2f
    Dim i As Long, j As Long, k As Long, tile As Long, decalNo As Long, skipDecalDraw As Boolean
    Dim dxdy(0 To 3, 0 To 1) As Long, groundsBytes(0 To 3) As Long, vs(0 To 4) As TerrainDecalVertex
    Dim dx As Long, dy As Long, grounds As Long, idx As Long, p0 As Long
    Dim e1 As Vector3f, e2 As Vector3f, uv1 As Vector2f, uv2 As Vector2f, r As Single
    Dim tangent As Vector3f, bitangent As Vector3f, n As Vector3f, t As Vector3f, w As Single

    dxdy(0, 0) = 0: dxdy(0, 1) = 0
    dxdy(1, 0) = 0: dxdy(1, 1) = 1
    dxdy(2, 0) = 1: dxdy(2, 1) = 1
    dxdy(3, 0) = 1: dxdy(3, 1) = 0

    For i = x * sectorSize To x * sectorSize + sectorSize - 1
        For j = y * sectorSize To y * sectorSize + sectorSize - 1
            If i < 0 Or j < 0 Or i >= mapWidth Or j >= mapHeight Then
                GoTo continue_
            End If

            tile = mapTile(i, j)
            center = getTileTexArrCoords(uv, psMapTiles(tile).texture)
            decalNo = TileNumber_tile(psMapTiles(tile).texture)
            skipDecalDraw = Not TILE_HAS_DECAL(tile)
            If terrainShaderQuality = CLASSIC Then
                ' in Classic mode, skip drawing any that are the "water only" decal (water is handled separately as a prior pass)
                If Not skipDecalDraw Then
                    If isOnlyWater(i, j) Then skipDecalDraw = (decalNo = 17) ' Magic number hack, but decal # 17 is always the *water only* tile in the legacy terrain tilesets we ship // TODO: Figure out a better way of determining this from the tileset data?
                End If
            End If
            If skipDecalDraw Then
                decalNo = -1
            End If

            For k = 0 To 3
                dx = dxdy(k, 0)
                dy = dxdy(k, 1)
                getGridPos pos, i + dx, j + dy, False, False
                vs(k).pos = glm_vec3(pos.x, pos.y, pos.z)
                vs(k).decalUv = uv(dx * 2 + dy)
                vs(k).normal = getGridNormal(i + dx, j + dy)
                vs(k).decalNo = decalNo
                groundsBytes(k) = psMapTiles(mapTile(i + dx, j + dy)).ground
                vs(k).groundWeights = 0                         ' clear()
                Select Case k                                   ' setByte(k, 255)
                Case 0: vs(k).groundWeights = &HFF&
                Case 1: vs(k).groundWeights = &HFF00&
                Case 2: vs(k).groundWeights = &HFF0000
                Case 3: vs(k).groundWeights = &HFF000000
                End Select
            Next
            grounds = pal_RGBA(groundsBytes(0), groundsBytes(1), groundsBytes(2), groundsBytes(3))
            ' 4 = center;
            getGridPos pos, i, j, True, False
            vs(4).pos = glm_vec3(pos.x, pos.y, pos.z)
            vs(4).decalUv = center
            vs(4).normal = getGridNormal(i, j, True)
            vs(4).decalNo = decalNo
            vs(4).groundWeights = 0 ' special value for shader.
            For k = 0 To 4
                vs(k).grounds = grounds
            Next

            terrainDecalData(terrainDecalSize) = vs(4): terrainDecalSize = terrainDecalSize + 1
            terrainDecalData(terrainDecalSize) = vs(1): terrainDecalSize = terrainDecalSize + 1
            terrainDecalData(terrainDecalSize) = vs(0): terrainDecalSize = terrainDecalSize + 1

            terrainDecalData(terrainDecalSize) = vs(4): terrainDecalSize = terrainDecalSize + 1
            terrainDecalData(terrainDecalSize) = vs(2): terrainDecalSize = terrainDecalSize + 1
            terrainDecalData(terrainDecalSize) = vs(1): terrainDecalSize = terrainDecalSize + 1

            terrainDecalData(terrainDecalSize) = vs(4): terrainDecalSize = terrainDecalSize + 1
            terrainDecalData(terrainDecalSize) = vs(3): terrainDecalSize = terrainDecalSize + 1
            terrainDecalData(terrainDecalSize) = vs(2): terrainDecalSize = terrainDecalSize + 1

            terrainDecalData(terrainDecalSize) = vs(4): terrainDecalSize = terrainDecalSize + 1
            terrainDecalData(terrainDecalSize) = vs(0): terrainDecalSize = terrainDecalSize + 1
            terrainDecalData(terrainDecalSize) = vs(3): terrainDecalSize = terrainDecalSize + 1

            ' calc tangents
            For idx = terrainDecalSize - 3 * 4 To terrainDecalSize - 1 Step 3
                p0 = idx
                e1 = glm_vec3_sub(terrainDecalData(p0 + 1).pos, terrainDecalData(p0).pos)
                e2 = glm_vec3_sub(terrainDecalData(p0 + 2).pos, terrainDecalData(p0).pos)
                uv1.x = terrainDecalData(p0 + 1).decalUv.x - terrainDecalData(p0).decalUv.x
                uv1.y = terrainDecalData(p0 + 1).decalUv.y - terrainDecalData(p0).decalUv.y
                uv2.x = terrainDecalData(p0 + 2).decalUv.x - terrainDecalData(p0).decalUv.x
                uv2.y = terrainDecalData(p0 + 2).decalUv.y - terrainDecalData(p0).decalUv.y
                r = fdiv(1!, uv1.x * uv2.y - uv2.x * uv1.y)
                tangent = glm_normalize3(glm_vec3_mul(glm_vec3_sub(glm_vec3_mul(e1, uv2.y), glm_vec3_mul(e2, uv1.y)), r))
                bitangent = glm_normalize3(glm_vec3_mul(glm_vec3_add(glm_vec3_mul(e1, -uv2.x), glm_vec3_mul(e2, uv1.x)), r))
                For k = 0 To 2
                    n = terrainDecalData(p0 + k).normal
                    t = glm_normalize3(glm_vec3_sub(tangent, glm_vec3_mul(n, glm_dot3(tangent, n))))
                    w = 1! ' not mirrored
                    If glm_dot3(glm_cross(n, t), bitangent) < 0! Then
                        w = -1! ' we're mirrored
                    End If
                    terrainDecalData(p0 + k).decalTangent = glm_vec4_v3(t, w)
                Next
            Next
continue_:
        Next
    Next
End Sub

' Update the sector for when the terrain is changed.
Private Sub updateSectorGeometry(ByVal x As Long, ByVal y As Long)
    Dim geometry() As TerrainVertex, water() As Vector4f, s As Long
    Dim geometrySize As Long, waterSize As Long, terrainDecalSize As Long
    geometrySize = 0
    waterSize = 0
    s = x * ySectors + y

    ReDim geometry(0 To sectors(s).geometrySize)
    ReDim water(0 To sectors(s).waterSize)

    setSectorGeometry x, y, geometry, water, geometrySize, waterSize
    If Not (geometrySize = sectors(s).geometrySize) Then ASSERT_FAILURE "geometrySize == sectors[x * ySectors + y].geometrySize", "src/terrain.cpp:539", "updateSectorGeometry", "something went seriously wrong updating the terrain"
    If Not (waterSize = sectors(s).waterSize) Then ASSERT_FAILURE "waterSize == sectors[x * ySectors + y].waterSize", "src/terrain.cpp:540", "updateSectorGeometry", "something went seriously wrong updating the terrain"

    geometryVBO.update LenB(geometry(0)) * sectors(s).geometryOffset, _
                       LenB(geometry(0)) * sectors(s).geometrySize, VarPtr(geometry(0)), _
                       update_flag_non_overlapping_updates_promise
    waterVBO.update LenB(water(0)) * sectors(s).waterOffset, _
                    LenB(water(0)) * sectors(s).waterSize, VarPtr(water(0)), _
                    update_flag_non_overlapping_updates_promise

    Erase geometry
    Erase water

    ' reuse a buffer to avoid repeated allocations if possible
    If sectors(s).terrainAndDecalSize > terrainDecalVertexUpdateBuffer_size Then
        ReDim terrainDecalVertexUpdateBuffer(0 To sectors(s).terrainAndDecalSize - 1)
        terrainDecalVertexUpdateBuffer_size = sectors(s).terrainAndDecalSize
    End If
    terrainDecalSize = 0
    setSectorDecalVertex_SinglePass x, y, terrainDecalVertexUpdateBuffer, terrainDecalSize
    If Not (terrainDecalSize = sectors(s).terrainAndDecalSize) Then ASSERT_FAILURE "terrainDecalSize == sectors[x * ySectors + y].terrainAndDecalSize", "src/terrain.cpp:555", "updateSectorGeometry", "Sizes don't match!"
    If sectors(s).terrainAndDecalSize > 0 Then
        terrainDecalVBO.update LenB(terrainDecalVertexUpdateBuffer(0)) * sectors(s).terrainAndDecalOffset, _
                               LenB(terrainDecalVertexUpdateBuffer(0)) * sectors(s).terrainAndDecalSize, VarPtr(terrainDecalVertexUpdateBuffer(0)), _
                               update_flag_non_overlapping_updates_promise
    End If
End Sub

' Mark all tiles that are influenced by this grid point as dirty.
' Dirty sectors will later get updated by updateSectorGeometry.
Public Sub markTileDirty(ByVal i As Long, ByVal j As Long)
    Dim x As Long, y As Long

    If Not terrainInitialised Then
        Exit Sub ' will be updated anyway
    End If

    x = i \ sectorSize
    y = j \ sectorSize
    If x < xSectors And y < ySectors Then ' could be on the lower or left edge of the map
        sectors(x * ySectors + y).dirty = True
    End If

    ' it could be on an edge, so update for all sectors it is in
    If x * sectorSize = i And x > 0 Then
        If x - 1 < xSectors And y < ySectors Then
            sectors((x - 1) * ySectors + y).dirty = True
        End If
    End If
    If y * sectorSize = j And y > 0 Then
        If x < xSectors And y - 1 < ySectors Then
            sectors(x * ySectors + (y - 1)).dirty = True
        End If
    End If
    If x * sectorSize = i And x > 0 And y * sectorSize = j And y > 0 Then
        If x - 1 < xSectors And y - 1 < ySectors Then
            sectors((x - 1) * ySectors + (y - 1)).dirty = True
        End If
    End If
End Sub

' Mark all tiles dirty
' (when switching between classic and other modes, recalculating tile heights is required due to water)
Public Sub dirtyAllSectors()
    Dim i As Long, len_ As Long
    If sectors_allocated Then
        len_ = xSectors * ySectors
        For i = 0 To len_ - 1
            sectors(i).dirty = True
        Next
    End If
End Sub

Public Function setTerrainMappingTexturesMaxSize(ByVal texSize As Long) As Boolean
    Dim isPowerOfTwo As Boolean
    isPowerOfTwo = Not (texSize = 0) And (texSize And (texSize - 1)) = 0
    If Not isPowerOfTwo Or texSize < MIN_TERRAIN_TEXTURE_SIZE Or texSize > 1024 Then
        If texSize <> 0 Then
            debug_ 628, LOG_ERROR, "setTerrainMappingTexturesMaxSize", "Attempted to set bad terrain mapping texture max size %d! Ignored.", texSize
        End If
        setTerrainMappingTexturesMaxSize = False
        Exit Function
    End If
    maxTerrainMappingTextureSize = texSize
    debug_ 633, LOG_TEXTURE, "setTerrainMappingTexturesMaxSize", "terrain mapping texture max size set to %d", texSize
    setTerrainMappingTexturesMaxSize = True
End Function

Public Function getTerrainMappingTexturesMaxSize() As Long
    getTerrainMappingTexturesMaxSize = maxTerrainMappingTextureSize
End Function

Public Function getWaterClassicTexture() As gfx_api_texture
    Dim maxGfxTextureSize As Long, maxTerrainTextureSize As Long, legacyWaterDecalPath As String
    If Not waterClassicTexture Is Nothing Then
        Set getWaterClassicTexture = waterClassicTexture
        Exit Function
    End If

    maxGfxTextureSize = gfx_api_context_get_context_value(context_value_MAX_TEXTURE_SIZE)
    maxTerrainTextureSize = getTextureSize()
    If maxGfxTextureSize < maxTerrainTextureSize Then maxTerrainTextureSize = maxGfxTextureSize
    If maxTerrainTextureSize < MIN_TERRAIN_TEXTURE_SIZE Then maxTerrainTextureSize = MIN_TERRAIN_TEXTURE_SIZE

    legacyWaterDecalPath = tilesetDir & "-" & CStr(getCurrentTileTextureSize()) & "/tile-17.png" ' TODO: This is currently hard-coded for legacy tileset textures...
    Set waterClassicTexture = gfx_api_context_loadTextureFromFile(legacyWaterDecalPath, texture_type_game_texture, maxTerrainTextureSize, maxTerrainTextureSize)
    Set getWaterClassicTexture = waterClassicTexture
End Function

' (VB) loadWaterTextures' lambda checkTex
Private Function loadWaterTextures_checkTex(ByVal maxTerrainTextureSize As Long, ByVal fileName As String, ByVal typ As texture_type) As gfx_api_texture
    Dim fullName As String, imageLoadFilename As String
    fullName = "texpages/" & fileName
    imageLoadFilename = gfx_api_imageLoadFilenameFromInputFilename(fullName)
    If PHYSFS_exists(imageLoadFilename) = 0 Then
        Set loadWaterTextures_checkTex = Nothing
        Exit Function
    End If
    Set loadWaterTextures_checkTex = gfx_api_context_loadTextureFromFile(imageLoadFilename, typ, maxTerrainTextureSize, maxTerrainTextureSize)
End Function

' (VB) loadWaterTextures' lambda optTexturenameToPath
Private Function loadWaterTextures_optTexturenameToPath(ByVal textureFilename As String) As String
    Dim fullName As String, imageLoadFilename As String
    If Len(textureFilename) = 0 Then
        Exit Function
    End If
    fullName = "texpages/" & textureFilename
    imageLoadFilename = gfx_api_imageLoadFilenameFromInputFilename(fullName)
    If PHYSFS_exists(imageLoadFilename) = 0 Then
        Exit Function
    End If
    loadWaterTextures_optTexturenameToPath = fullName
End Function

Public Sub loadWaterTextures(ByVal maxTerrainTextureSize As Long, Optional ByVal has_maxTerrainAuxTextureSizeOpt As Boolean = False, Optional ByVal maxTerrainAuxTextureSizeOpt As Long = 0)
    Dim waterTextureFilenames(0 To 1) As String, waterTextureFilenames_nm(0 To 1) As String, waterTextureFilenames_sm(0 To 1) As String
    Dim i As Long, allPresent As Boolean, anyPresent As Boolean, maxTerrainAuxTextureSize As Long, pWaterTexClassic As gfx_api_texture

    Set waterTexturesNormal.tex1 = Nothing: Set waterTexturesNormal.tex2 = Nothing                                     ' clear()
    Set waterTexturesHigh.tex = Nothing: Set waterTexturesHigh.tex_nm = Nothing: Set waterTexturesHigh.tex_sm = Nothing ' clear()

    If terrainShaderQuality = MEDIUM Then
        Set waterTexturesNormal.tex1 = loadWaterTextures_checkTex(maxTerrainTextureSize, "page-80-water-1.png", texture_type_game_texture)
        Set waterTexturesNormal.tex2 = loadWaterTextures_checkTex(maxTerrainTextureSize, "page-81-water-2.png", texture_type_specular_map)
    ElseIf terrainShaderQuality = NORMAL_MAPPING Then
        ' check water optional textures
        waterTextureFilenames(0) = loadWaterTextures_optTexturenameToPath("page-80-water-1.png")
        waterTextureFilenames(1) = loadWaterTextures_optTexturenameToPath("page-81-water-2.png")
        waterTextureFilenames_nm(0) = loadWaterTextures_optTexturenameToPath("page-80-water-1_nm.png")
        waterTextureFilenames_nm(1) = loadWaterTextures_optTexturenameToPath("page-81-water-2_nm.png")
        waterTextureFilenames_sm(0) = loadWaterTextures_optTexturenameToPath("page-80-water-1_sm.png")
        waterTextureFilenames_sm(1) = loadWaterTextures_optTexturenameToPath("page-81-water-2_sm.png")

        allPresent = True
        For i = 0 To 1
            If Len(waterTextureFilenames(i)) = 0 Then allPresent = False
        Next
        If Not allPresent Then
            debug_ 744, LOG_FATAL, "loadWaterTextures", "Missing one or more base water textures?"
            Exit Sub
        End If

        Set waterTexturesHigh.tex = gfx_api_context_loadTextureArrayFromFiles(waterTextureFilenames, 2, texture_type_game_texture, maxTerrainTextureSize, maxTerrainTextureSize, GenerateDefaultTexture_none, ProgressCallback_resDoResLoadCallback)
        Set waterTexturesHigh.tex_nm = Nothing
        Set waterTexturesHigh.tex_sm = Nothing

        If has_maxTerrainAuxTextureSizeOpt Then maxTerrainAuxTextureSize = maxTerrainAuxTextureSizeOpt Else maxTerrainAuxTextureSize = maxTerrainTextureSize

        anyPresent = False
        For i = 0 To 1
            If Len(waterTextureFilenames_nm(i)) <> 0 Then anyPresent = True
        Next
        If anyPresent Then
            Set waterTexturesHigh.tex_nm = gfx_api_context_loadTextureArrayFromFiles(waterTextureFilenames_nm, 2, texture_type_normal_map, maxTerrainAuxTextureSize, maxTerrainAuxTextureSize, GenerateDefaultTexture_normalMap, ProgressCallback_resDoResLoadCallback)
            If Not (Not waterTexturesHigh.tex_nm Is Nothing) Then ASSERT_FAILURE "waterTexturesHigh.tex_nm != nullptr", "src/terrain.cpp:774", "loadWaterTextures", "Failed to load water normals"
        End If
        anyPresent = False
        For i = 0 To 1
            If Len(waterTextureFilenames_sm(i)) <> 0 Then anyPresent = True
        Next
        If anyPresent Then
            Set waterTexturesHigh.tex_sm = gfx_api_context_loadTextureArrayFromFiles(waterTextureFilenames_sm, 2, texture_type_specular_map, maxTerrainAuxTextureSize, maxTerrainAuxTextureSize, GenerateDefaultTexture_zeroed, ProgressCallback_resDoResLoadCallback)
            If Not (Not waterTexturesHigh.tex_sm Is Nothing) Then ASSERT_FAILURE "waterTexturesHigh.tex_sm != nullptr", "src/terrain.cpp:786", "loadWaterTextures", "Failed to load water specular maps"
        End If
    ElseIf terrainShaderQuality = CLASSIC Then
        ' preload classic water texture
        Set pWaterTexClassic = getWaterClassicTexture()
        If Not (Not pWaterTexClassic Is Nothing) Then ASSERT_FAILURE "pWaterTexClassic != nullptr", "src/terrain.cpp:793", "loadWaterTextures", "Failed to load classic water texture?": Exit Sub
    Else
        ASSERT_FAILURE "false", "src/terrain.cpp:797", "loadWaterTextures", "Unexpected terrainShaderQuality: %u", terrainShaderQuality
        Exit Sub
    End If
End Sub

' (VB) loadTerrainTextures_SinglePass' lambda optTexturenameToPath
Private Function optTexturenameToPath(ByVal textureFilename As String) As String
    If Len(textureFilename) = 0 Then
        Exit Function
    End If
    optTexturenameToPath = "texpages/" & textureFilename
End Function

' (VB) std::any_of(names, !isEmpty)
Private Function anyNotEmpty(names() As String, ByVal count As Long) As Boolean
    Dim i As Long
    For i = 0 To count - 1
        If Len(names(i)) <> 0 Then
            anyNotEmpty = True
            Exit Function
        End If
    Next
End Function

Public Sub loadTerrainTextures_SinglePass(ByVal mapTileset As MAP_TILESET)
    Dim maxGfxTextureSize As Long, maxTerrainTextureSize As Long, maxTerrainAuxTextureSize As Long
    Dim groundTextureFilenames() As String, groundTextureFilenames_nm() As String
    Dim groundTextureFilenames_spec() As String, groundTextureFilenames_height() As String
    Dim layer As Long, n As Long, groundType As GROUND_TYPE

    If Not (getNumGroundTypes() <> 0) Then ASSERT_FAILURE "getNumGroundTypes()", "src/terrain.cpp:803", "loadTerrainTextures_SinglePass", "Ground type was not set, no textures will be seen.": Exit Sub

    maxGfxTextureSize = gfx_api_context_get_context_value(context_value_MAX_TEXTURE_SIZE)
    maxTerrainTextureSize = getTextureSize()
    If maxGfxTextureSize < maxTerrainTextureSize Then maxTerrainTextureSize = maxGfxTextureSize
    If maxTerrainTextureSize < MIN_TERRAIN_TEXTURE_SIZE Then maxTerrainTextureSize = MIN_TERRAIN_TEXTURE_SIZE

    Set groundTexArr = Nothing
    Set groundNormalArr = Nothing
    Set groundSpecularArr = Nothing
    Set groundHeightArr = Nothing
    Set waterClassicTexture = Nothing

    n = getNumGroundTypes()
    ReDim groundTextureFilenames(0 To n - 1)
    ReDim groundTextureFilenames_nm(0 To n - 1)
    ReDim groundTextureFilenames_spec(0 To n - 1)
    ReDim groundTextureFilenames_height(0 To n - 1)

    ' for each terrain layer
    For layer = 0 To n - 1
        groundType = getGroundType(layer)
        groundTextureFilenames(layer) = optTexturenameToPath(groundType.textureName)
        groundTextureFilenames_nm(layer) = optTexturenameToPath(groundType.normalMapTextureName)
        groundTextureFilenames_spec(layer) = optTexturenameToPath(groundType.specularMapTextureName)
        groundTextureFilenames_height(layer) = optTexturenameToPath(groundType.heightMapTextureName)
    Next

    ' load the textures into the texture arrays
    Set groundTexArr = gfx_api_context_loadTextureArrayFromFiles(groundTextureFilenames, n, texture_type_game_texture, maxTerrainTextureSize, maxTerrainTextureSize, GenerateDefaultTexture_none, ProgressCallback_resDoResLoadCallback)
    If Not (Not groundTexArr Is Nothing) Then ASSERT_FAILURE "groundTexArr != nullptr", "src/terrain.cpp:839", "loadTerrainTextures_SinglePass", "Failed to load terrain textures"

    maxTerrainAuxTextureSize = getTextureSize()
    If getTerrainMappingTexturesMaxSize() < maxTerrainAuxTextureSize Then maxTerrainAuxTextureSize = getTerrainMappingTexturesMaxSize()
    If maxGfxTextureSize < maxTerrainAuxTextureSize Then maxTerrainAuxTextureSize = maxGfxTextureSize
    If maxTerrainAuxTextureSize < MIN_TERRAIN_TEXTURE_SIZE Then maxTerrainAuxTextureSize = MIN_TERRAIN_TEXTURE_SIZE

    If terrainShaderQuality = NORMAL_MAPPING Then
        If anyNotEmpty(groundTextureFilenames_nm, n) Then
            Set groundNormalArr = gfx_api_context_loadTextureArrayFromFiles(groundTextureFilenames_nm, n, texture_type_normal_map, maxTerrainAuxTextureSize, maxTerrainAuxTextureSize, GenerateDefaultTexture_normalMap, ProgressCallback_resDoResLoadCallback)
            If Not (Not groundNormalArr Is Nothing) Then ASSERT_FAILURE "groundNormalArr != nullptr", "src/terrain.cpp:865", "loadTerrainTextures_SinglePass", "Failed to load terrain normals"
        End If
        If anyNotEmpty(groundTextureFilenames_spec, n) Then
            Set groundSpecularArr = gfx_api_context_loadTextureArrayFromFiles(groundTextureFilenames_spec, n, texture_type_specular_map, maxTerrainAuxTextureSize, maxTerrainAuxTextureSize, GenerateDefaultTexture_zeroed, ProgressCallback_resDoResLoadCallback)
            If Not (Not groundSpecularArr Is Nothing) Then ASSERT_FAILURE "groundSpecularArr != nullptr", "src/terrain.cpp:877", "loadTerrainTextures_SinglePass", "Failed to load terrain specular maps"
        End If
        If anyNotEmpty(groundTextureFilenames_height, n) Then
            Set groundHeightArr = gfx_api_context_loadTextureArrayFromFiles(groundTextureFilenames_height, n, texture_type_height_map, maxTerrainAuxTextureSize, maxTerrainAuxTextureSize, GenerateDefaultTexture_zeroed, ProgressCallback_resDoResLoadCallback)
            If Not (Not groundHeightArr Is Nothing) Then ASSERT_FAILURE "groundHeightArr != nullptr", "src/terrain.cpp:889", "loadTerrainTextures_SinglePass", "Failed to load terrain height maps"
        End If
    End If

    ' load water textures
    loadWaterTextures maxTerrainTextureSize, True, maxTerrainAuxTextureSize
End Sub

Public Sub loadTerrainTextures(ByVal mapTileset As MAP_TILESET)
    ' (VB) no graphics context (the --vbtest runs): nothing to load
    If Not gfx_api_context_isInitialized() Then Exit Sub
    If Not (getNumGroundTypes() <> 0) Then ASSERT_FAILURE "getNumGroundTypes()", "src/terrain.cpp:899", "loadTerrainTextures", "Ground type was not set, no textures will be seen.": Exit Sub
    loadTerrainTextures_SinglePass mapTileset
End Sub

Public Sub reloadTerrainTextures()
    If getNumGroundTypes() = 0 Then
        Exit Sub ' nothing loaded yet
    End If

    loadTerrainTextures currentMapTileset
End Sub

' (VB) the q(i,j,center) macro of initTerrain
Private Function q(ByVal x As Long, ByVal y As Long, ByVal i As Long, ByVal j As Long, ByVal center As Long) As Long
    q = (x * ySectors + y) * (sectorSize + 1) * (sectorSize + 1) * 2 + (i * (sectorSize + 1) + j) * 2 + center
End Function

' Check what the videocard + drivers support and divide the loaded map into sectors that can be drawn.
' It also determines the lightmap size.
Public Function initTerrain() As Boolean
    Dim i As Long, j As Long, x As Long, y As Long, s As Long

    Dim geometry() As TerrainVertex
    Dim water() As Vector4f
    Dim geometrySize As Long, geometryIndexSize As Long
    Dim waterSize As Long, waterIndexSize As Long
    Dim geometryIndex() As Long
    Dim waterIndex() As Long
    Dim maxSectorSizeIndices As Long, maxSectorSizeVertices As Long
    Dim decreasedSize As Boolean
    Dim vertSize As Long, terrainDecalData() As TerrainDecalVertex, terrainDecalSize As Long
    Dim lightmapChannels As Long, cols As mat4

    decreasedSize = False

    ' this information is useful to prevent crashes with buggy opengl implementations
    GLmaxElementsVertices = gfx_api_context_get_context_value(context_value_MAX_ELEMENTS_VERTICES)
    GLmaxElementsIndices = gfx_api_context_get_context_value(context_value_MAX_ELEMENTS_INDICES)

    ' testing for crappy cards
    debug_ 935, LOG_TERRAIN, "initTerrain", "GL_MAX_ELEMENTS_VERTICES: %i", GLmaxElementsVertices
    debug_ 936, LOG_TERRAIN, "initTerrain", "GL_MAX_ELEMENTS_INDICES:  %i", GLmaxElementsIndices

    ' now we know these values, determine the maximum sector size achievable
    maxSectorSizeVertices = iSqrt(GLmaxElementsVertices \ 2) - 1
    maxSectorSizeIndices = iSqrt(GLmaxElementsIndices \ 12)

    debug_ 942, LOG_TERRAIN, "initTerrain", "preferred sector size: %i", sectorSize
    debug_ 943, LOG_TERRAIN, "initTerrain", "maximum sector size due to vertices: %i", maxSectorSizeVertices
    debug_ 944, LOG_TERRAIN, "initTerrain", "maximum sector size due to indices: %i", maxSectorSizeIndices

    If sectorSize > maxSectorSizeVertices Then
        sectorSize = maxSectorSizeVertices
        decreasedSize = True
    End If
    If sectorSize > maxSectorSizeIndices Then
        sectorSize = maxSectorSizeIndices
        decreasedSize = True
    End If
    If decreasedSize Then
        If sectorSize < 1 Then
            debug_ 960, LOG_WARNING, "initTerrain", "GL_MAX_ELEMENTS_VERTICES: %i", GLmaxElementsVertices
            debug_ 961, LOG_WARNING, "initTerrain", "GL_MAX_ELEMENTS_INDICES:  %i", GLmaxElementsIndices
            debug_ 962, LOG_WARNING, "initTerrain", "maximum sector size due to vertices: %i", maxSectorSizeVertices
            debug_ 963, LOG_WARNING, "initTerrain", "maximum sector size due to indices: %i", maxSectorSizeIndices
            debug_ 964, LOG_ERROR, "initTerrain", "Your graphics card and/or drivers do not seem to support glDrawRangeElements, needed for the terrain renderer."
            debug_ 965, LOG_ERROR, "initTerrain", "- Do other 3D games work?"
            debug_ 966, LOG_ERROR, "initTerrain", "- Did you install the latest drivers correctly?"
            debug_ 967, LOG_ERROR, "initTerrain", "- Do you have a 3D window manager (Aero/Compiz) running?"
            initTerrain = False
            Exit Function
        End If
        debug_ 970, LOG_WARNING, "initTerrain", "decreasing sector size to %i to fit graphics card constraints", sectorSize
    End If

    ' +4 = +1 for iHypot rounding, +1 for sector size rounding, +2 for edge of visibility
    terrainDistance = iHypot(visibleTiles.x \ 2, visibleTiles.y \ 2) + 4 + sectorSize \ 2
    debug_ 975, LOG_TERRAIN, "initTerrain", "visible tiles x:%i y: %i", visibleTiles.x, visibleTiles.y
    debug_ 976, LOG_TERRAIN, "initTerrain", "terrain view distance: %i", terrainDistance

    '////////////////////
    ' Create the sectors
    xSectors = (mapWidth + sectorSize - 1) \ sectorSize
    ySectors = (mapHeight + sectorSize - 1) \ sectorSize
    ReDim sectors(0 To xSectors * ySectors)
    sectors_allocated = True

    '//////////////////
    ' fill the geometry part of the sectors
    vertSize = xSectors * ySectors * (sectorSize + 1) * (sectorSize + 1) * 2
    ReDim geometry(0 To vertSize)
    ReDim geometryIndex(0 To xSectors * ySectors * sectorSize * sectorSize * 12)
    geometrySize = 0
    geometryIndexSize = 0

    ReDim water(0 To vertSize)
    ReDim waterIndex(0 To xSectors * ySectors * sectorSize * sectorSize * 12)
    waterSize = 0
    waterIndexSize = 0
    For x = 0 To xSectors - 1
        For y = 0 To ySectors - 1
            s = x * ySectors + y
            sectors(s).dirty = False
            sectors(s).geometryOffset = geometrySize
            sectors(s).geometrySize = 0
            sectors(s).waterOffset = waterSize
            sectors(s).waterSize = 0

            setSectorGeometry x, y, geometry, water, geometrySize, waterSize

            sectors(s).geometrySize = geometrySize - sectors(s).geometryOffset
            sectors(s).waterSize = waterSize - sectors(s).waterOffset
            ' and do the index buffers
            sectors(s).geometryIndexOffset = geometryIndexSize
            sectors(s).geometryIndexSize = 0
            sectors(s).waterIndexOffset = waterIndexSize
            sectors(s).waterIndexSize = 0

            For i = 0 To sectorSize - 1
                For j = 0 To sectorSize - 1
                    If x * sectorSize + i >= mapWidth Or y * sectorSize + j >= mapHeight Then
                        GoTo continue_ ' off map, so skip
                    End If

                    ' One tile is composed of 4 triangles,
                    ' we need _2_ vertices per tile (1)
                    '      e.g. center and bottom left
                    '  the other 3 vertices are from the adjacent tiles
                    '  on their top and right.
                    ' (1) The top row and right column of tiles need 4 vertices per tile
                    '  because they do not have adjacent tiles on their top and right,
                    '  that is why we add _1_ row and _1_ column to provide the geometry
                    '  for these tiles.
                    ' This is the source of the '*2' and '+1' in the index math below.
                    ' First triangle
                    geometryIndex(geometryIndexSize + 0) = q(x, y, i, j, 1)      ' Center vertex
                    geometryIndex(geometryIndexSize + 1) = q(x, y, i, j, 0)      ' Bottom left
                    geometryIndex(geometryIndexSize + 2) = q(x, y, i + 1, j, 0)  ' Bottom right
                    ' Second triangle
                    geometryIndex(geometryIndexSize + 3) = q(x, y, i, j, 1)      ' Center vertex
                    geometryIndex(geometryIndexSize + 4) = q(x, y, i, j + 1, 0)  ' Top left
                    geometryIndex(geometryIndexSize + 5) = q(x, y, i, j, 0)      ' Bottom left
                    ' Third triangle
                    geometryIndex(geometryIndexSize + 6) = q(x, y, i, j, 1)      ' Center vertex
                    geometryIndex(geometryIndexSize + 7) = q(x, y, i + 1, j + 1, 0) ' Top right
                    geometryIndex(geometryIndexSize + 8) = q(x, y, i, j + 1, 0)  ' Top left
                    ' Fourth triangle
                    geometryIndex(geometryIndexSize + 9) = q(x, y, i, j, 1)      ' Center vertex
                    geometryIndex(geometryIndexSize + 10) = q(x, y, i + 1, j, 0) ' Bottom right
                    geometryIndex(geometryIndexSize + 11) = q(x, y, i + 1, j + 1, 0) ' Top right
                    geometryIndexSize = geometryIndexSize + 12
                    If isWater(i + x * sectorSize, j + y * sectorSize) Then
                        waterIndex(waterIndexSize + 0) = q(x, y, i, j, 1)
                        waterIndex(waterIndexSize + 1) = q(x, y, i, j, 0)
                        waterIndex(waterIndexSize + 2) = q(x, y, i + 1, j, 0)

                        waterIndex(waterIndexSize + 3) = q(x, y, i, j, 1)
                        waterIndex(waterIndexSize + 4) = q(x, y, i, j + 1, 0)
                        waterIndex(waterIndexSize + 5) = q(x, y, i, j, 0)

                        waterIndex(waterIndexSize + 6) = q(x, y, i, j, 1)
                        waterIndex(waterIndexSize + 7) = q(x, y, i + 1, j + 1, 0)
                        waterIndex(waterIndexSize + 8) = q(x, y, i, j + 1, 0)

                        waterIndex(waterIndexSize + 9) = q(x, y, i, j, 1)
                        waterIndex(waterIndexSize + 10) = q(x, y, i + 1, j, 0)
                        waterIndex(waterIndexSize + 11) = q(x, y, i + 1, j + 1, 0)
                        waterIndexSize = waterIndexSize + 12
                    End If
continue_:
                Next
            Next
            sectors(s).geometryIndexSize = geometryIndexSize - sectors(s).geometryIndexOffset
            sectors(s).waterIndexSize = waterIndexSize - sectors(s).waterIndexOffset
        Next
    Next
    Set geometryVBO = Nothing
    Set geometryVBO = gfx_api_context_create_buffer_object(buffer_usage_vertex_buffer, buffer_storage_hint_dynamic_draw, "terrain::geometryVBO")
    geometryVBO.upload LenB(geometry(0)) * geometrySize, VarPtr(geometry(0))
    Erase geometry

    Set geometryIndexVBO = Nothing
    Set geometryIndexVBO = gfx_api_context_create_buffer_object(buffer_usage_index_buffer, buffer_storage_hint_static_draw, "terrain::geometryIndexVBO")
    geometryIndexVBO.upload 4 * geometryIndexSize, VarPtr(geometryIndex(0))   ' sizeof(GLuint)*geometryIndexSize
    Erase geometryIndex

    Set waterVBO = Nothing
    Set waterVBO = gfx_api_context_create_buffer_object(buffer_usage_vertex_buffer, buffer_storage_hint_dynamic_draw, "terrain::waterVBO")
    waterVBO.upload LenB(water(0)) * waterSize, VarPtr(water(0))
    Erase water

    Set waterIndexVBO = Nothing
    If waterIndexSize > 0 Then
        Set waterIndexVBO = gfx_api_context_create_buffer_object(buffer_usage_index_buffer, buffer_storage_hint_static_draw, "terrain::waterIndexVBO")
        waterIndexVBO.upload 4 * waterIndexSize, VarPtr(waterIndex(0))
    Else
        Set waterIndexVBO = Nothing
    End If
    Erase waterIndex

    ' and finally the decals
    ReDim terrainDecalData(0 To mapWidth * mapHeight * 12)
    terrainDecalSize = 0

    For x = 0 To xSectors - 1
        For y = 0 To ySectors - 1
            s = x * ySectors + y
            sectors(s).terrainAndDecalOffset = terrainDecalSize
            sectors(s).terrainAndDecalSize = 0
            setSectorDecalVertex_SinglePass x, y, terrainDecalData, terrainDecalSize
            sectors(s).terrainAndDecalSize = terrainDecalSize - sectors(s).terrainAndDecalOffset
        Next
    Next
    debug_ 1125, LOG_TERRAIN, "initTerrain", "%i decals found", terrainDecalSize \ 12

    Set terrainDecalVBO = Nothing
    If terrainDecalSize > 0 Then
        Set terrainDecalVBO = gfx_api_context_create_buffer_object(buffer_usage_vertex_buffer, buffer_storage_hint_dynamic_draw, "terrain::terrainDecalVBO")
        terrainDecalVBO.upload LenB(terrainDecalData(0)) * terrainDecalSize, VarPtr(terrainDecalData(0))
    Else
        Set terrainDecalVBO = Nothing
    End If
    Erase terrainDecalData

    lightmapLastUpdate = 0
    lightmapWidth = 1
    lightmapHeight = 1
    ' determine the smallest power-of-two size we can use for the lightmap
    Do
        lightmapWidth = lightmapWidth * 2
    Loop While mapWidth > lightmapWidth
    Do
        lightmapHeight = lightmapHeight * 2
    Loop While mapHeight > lightmapHeight
    debug_ 1150, LOG_TERRAIN, "initTerrain", "the size of the map is %ix%i", mapWidth, mapHeight
    debug_ 1151, LOG_TERRAIN, "initTerrain", "lightmap texture size is %zu x %zu", lightmapWidth, lightmapHeight

    lightmapValues.paramsXLight = glm_vec4(1! / world_coord(mapWidth) * (CSng(mapWidth) / CSng(lightmapWidth)), 0!, 0!, 0!)
    lightmapValues.paramsYLight = glm_vec4(0!, 0!, -1! / world_coord(mapHeight) * (CSng(mapHeight) / CSng(lightmapHeight)), 0!)

    ' shift the lightmap half a tile as lights are supposed to be placed at the center of a tile
    lightmapValues.lightMatrix = glm_translate(glm_vec3(1! / CSng(lightmapWidth) / 2!, 1! / CSng(lightmapHeight) / 2!, 0!))
    cols = glm_mat4_cols(lightmapValues.paramsXLight, lightmapValues.paramsYLight, glm_vec4(0!, 0!, 1!, 0!), glm_vec4(0!, 0!, 0!, 1!))
    lightmapValues.ModelUVLightmap = glm_mat4_mul(lightmapValues.lightMatrix, glm_transpose(cols))

    ' Prepare the lightmap pixmap and texture
    Set lightmapPixmap = New iV_Image

    lightmapChannels = 4 ' always use 4-channel (RGBA)
    If Not lightmapPixmap.allocate(lightmapWidth, lightmapHeight, lightmapChannels, True) Then
        debug_ 1166, LOG_FATAL, "initTerrain", "Out of memory!"
        exit_ 3                         ' abort()
        initTerrain = False
        Exit Function
    End If
    ReDim lightmapBytes(0 To lightmapWidth * lightmapHeight * lightmapChannels - 1)
    Set lightmap_texture = Nothing
    Set lightmap_texture = gfx_api_context_create_texture(1, lightmapPixmap.width(), lightmapPixmap.height(), lightmapPixmap.pixel_format_(), "mem::lightmap")

    lightmap_texture.upload 0, lightmapPixmap
    terrainInitialised = True

    initTerrain = True
End Function

' free all memory and opengl buffers used by the terrain renderer
Public Sub shutdownTerrain()
    Set geometryVBO = Nothing
    Set geometryIndexVBO = Nothing
    Set waterVBO = Nothing
    Set waterIndexVBO = Nothing
    Set terrainDecalVBO = Nothing

    Erase sectors                       ' sectors.reset()
    sectors_allocated = False

    Set lightmap_texture = Nothing
    Set lightmapPixmap = Nothing
    Erase lightmapBytes

    Set groundTexArr = Nothing
    Set groundNormalArr = Nothing
    Set groundSpecularArr = Nothing
    Set groundHeightArr = Nothing

    Set waterTexturesNormal.tex1 = Nothing: Set waterTexturesNormal.tex2 = Nothing                                     ' clear()
    Set waterTexturesHigh.tex = Nothing: Set waterTexturesHigh.tex_nm = Nothing: Set waterTexturesHigh.tex_sm = Nothing ' clear()
    Set waterClassicTexture = Nothing

    Set decalTexArr = Nothing
    Set decalNormalArr = Nothing
    Set decalSpecularArr = Nothing
    Set decalHeightArr = Nothing

    terrainInitialised = False
End Sub

' (VB) lightmap is getCurrentLightmapData() (WzGfx's LightMap: its pixels copied out at once)
Private Sub updateLightMap()
    Dim lightmapChannels As Long, i As Long, j As Long, psTile As Long, colour As Long, level As Long
    Dim r As Long, g As Long, b As Long, m As Long, p As Long, lm() As Long, n As Long
    Dim playerX As Single, playerY As Single, distA As Single, distB As Single, distC As Single, distD As Single
    Dim darken As Single, distToEdge As Single, fogOn As Boolean
    lightmapChannels = lightmapPixmap.channels() ' should always be 4 now...
    n = mapWidth * mapHeight
    If n <= 0 Then Exit Sub
    ReDim lm(0 To n - 1)
    CopyMemory lm(0), ByVal LightMap_dataAddress(), 4 * n
    fogOn = pie_GetFogStatus()
    playerX = map_coordf(playerPos.p.x)
    playerY = map_coordf(playerPos.p.z)
    For j = 0 To mapHeight - 1
        For i = 0 To mapWidth - 1
            psTile = mapTile(i, j)
            colour = lm(i + j * mapWidth)                   ' lightmap(i, j)
            level = CLng(Fix(psMapTiles(psTile).level)) And &HFF&

            r = colour And &HFF&
            g = (colour And &HFF00&) \ &H100&
            b = (colour And &HFF0000) \ &H10000
            If (psMapTiles(psTile).tileInfoBits And BITS_GATEWAY) <> 0 And showGateways Then
                g = 255
            End If
            If (psMapTiles(psTile).tileInfoBits And BITS_MARKED) <> 0 Then
                m = getModularScaledGraphicsTime(2048, 255)
                If m > 255 - m Then r = m And &HFF& Else r = (255 - m) And &HFF&
                If r \ 2 > level Then level = r \ 2
            End If

            p = (i + j * lightmapWidth) * lightmapChannels
            lightmapBytes(p + 0) = r
            lightmapBytes(p + 1) = g
            lightmapBytes(p + 2) = b
            ' store the "brightness" level in byte.a
            ' NOTE: This differs depending on whether using the single-pass terrain shader or the fallback terrain shaders
            ' (For more, see avUpdateTiles() and getTileIllumination())
            lightmapBytes(p + 3) = level

            If Not fogOn Then
                ' fade to black at the edges of the visible terrain area
                distA = i - (playerX - visibleTiles.x \ 2)
                distB = (playerX + visibleTiles.x \ 2) - i
                distC = j - (playerY - visibleTiles.y \ 2)
                distD = (playerY + visibleTiles.y \ 2) - j

                ' calculate the distance to the closest edge of the visible map
                ' determine the smallest distance
                distToEdge = distA
                If distB < distToEdge Then
                    distToEdge = distB
                End If
                If distC < distToEdge Then
                    distToEdge = distC
                End If
                If distD < distToEdge Then
                    distToEdge = distD
                End If

                darken = distToEdge / 2!
                If darken <= 0! Then
                    lightmapBytes(p + 0) = 0
                    lightmapBytes(p + 1) = 0
                    lightmapBytes(p + 2) = 0
                    lightmapBytes(p + 3) = 0
                ElseIf darken < 1! Then
                    lightmapBytes(p + 0) = CLng(Fix(lightmapBytes(p + 0) * darken))
                    lightmapBytes(p + 1) = CLng(Fix(lightmapBytes(p + 1) * darken))
                    lightmapBytes(p + 2) = CLng(Fix(lightmapBytes(p + 2) * darken))
                    lightmapBytes(p + 3) = CLng(Fix(lightmapBytes(p + 3) * darken))
                End If
            End If
        Next
    Next
    CopyMemory ByVal lightmapPixmap.bmp_w(), lightmapBytes(0), UBound(lightmapBytes) + 1
End Sub

Private Sub cullTerrain()
    Dim x As Long, y As Long, xPos As Single, yPos As Single, distance_ As Double, s As Long
    For x = 0 To xSectors - 1
        For y = 0 To ySectors - 1
            xPos = world_coord(x * sectorSize + sectorSize \ 2)
            yPos = world_coord(y * sectorSize + sectorSize \ 2)
            distance_ = (playerPos.p.x - xPos) ^ 2 + (playerPos.p.z - yPos) ^ 2
            s = x * ySectors + y

            If distance_ > CDbl(world_coord(terrainDistance)) ^ 2 Then
                sectors(s).draw = False
            Else
                sectors(s).draw = True
                If sectors(s).dirty Then
                    updateSectorGeometry x, y
                    sectors(s).dirty = False
                End If
            End If
        Next
    Next
End Sub

Private Sub drawDepthOnly(ModelViewProjection As mat4, paramsXLight As Vector4f, paramsYLight As Vector4f, ByVal withOffset As Boolean)
    Dim fogEnabled As Boolean, fogBegin As Single, fogEnd As Single, fogColour As Long
    Dim cb As constant_buffer_type_SHADER_TERRAIN_DEPTH, x As Long, y As Long, s As Long, pso As pipeline_state_helper
    renderState_get fogEnabled, fogBegin, fogEnd, fogColour

    ' bind the vertex buffer
    Set pso = gfx_api_pso_get(TerrainDepth)
    pso.bind
    pso.bind_textures lightmap_texture
    pso.bind_vertex_buffers geometryVBO
    cb.transform_matrix = ModelViewProjection
    cb.paramX = paramsXLight
    cb.paramY = paramsYLight
    cb.paramXLight = glm_vec4(0!, 0!, 0!, 0!)
    cb.paramYLight = glm_vec4(0!, 0!, 0!, 0!)
    cb.unused = glm_mat4(1!)
    cb.texture_matrix = glm_mat4(1!)
    cb.fog_colour = glm_vec4(0!, 0!, 0!, 0!)
    cb.fog_enabled = -CLng(fogEnabled)
    cb.fog_begin = fogBegin
    cb.fog_end = fogEnd
    cb.texture0 = 0
    cb.texture1 = 0
    pso.bind_constants VarPtr(cb), LenB(cb)
    gfx_api_context_bind_index_buffer geometryIndexVBO, index_type_u32

    If withOffset Then
        ' draw slightly higher distance than it actually is so it will not
        ' by accident obscure the actual terrain
        gfx_api_context_set_polygon_offset 0.1!, 1!
    End If

    For x = 0 To xSectors - 1
        For y = 0 To ySectors - 1
            s = x * ySectors + y
            If sectors(s).draw Then
                addDrawRangeElements TerrainDepth, _
                    sectors(s).geometryOffset, _
                    sectors(s).geometryOffset + sectors(s).geometrySize, _
                    sectors(s).geometryIndexSize, _
                    sectors(s).geometryIndexOffset
            End If
        Next
    Next
    finishDrawRangeElements TerrainDepth
    If withOffset Then
        gfx_api_context_set_polygon_offset 0!, 0!
    End If
    pso.unbind_vertex_buffers geometryVBO
    gfx_api_context_unbind_index_buffer geometryIndexVBO
End Sub

Private Sub drawDepthOnlyForDepthMap(ModelViewProjection As mat4, paramsXLight As Vector4f, paramsYLight As Vector4f, ByVal withOffset As Boolean)
    Dim fogEnabled As Boolean, fogBegin As Single, fogEnd As Single, fogColour As Long
    Dim cb As constant_buffer_type_SHADER_TERRAIN_DEPTHMAP, x As Long, y As Long, s As Long, pso As pipeline_state_helper
    renderState_get fogEnabled, fogBegin, fogEnd, fogColour

    ' bind the vertex buffer
    Set pso = gfx_api_pso_get(TerrainDepthOnlyForDepthMap)
    pso.bind
    pso.bind_textures lightmap_texture
    pso.bind_vertex_buffers geometryVBO
    cb.transform_matrix = ModelViewProjection
    cb.fog_enabled = -CLng(fogEnabled)
    cb.fog_begin = fogBegin
    cb.fog_end = fogEnd
    pso.bind_constants VarPtr(cb), LenB(cb)
    gfx_api_context_bind_index_buffer geometryIndexVBO, index_type_u32

'   if (withOffset)
'   {
'       // draw slightly higher distance than it actually is so it will not
'       // by accident obscure the actual terrain
'       gfx_api::context::get().set_polygon_offset(0.1f, 1.f);
'   }

    For x = 0 To xSectors - 1
        For y = 0 To ySectors - 1
            s = x * ySectors + y
            If sectors(s).draw Then
                addDrawRangeElements TerrainDepthOnlyForDepthMap, _
                    sectors(s).geometryOffset, _
                    sectors(s).geometryOffset + sectors(s).geometrySize, _
                    sectors(s).geometryIndexSize, _
                    sectors(s).geometryIndexOffset
            End If
        Next
    Next
    finishDrawRangeElements TerrainDepthOnlyForDepthMap
'   if (withOffset)
'   {
'       gfx_api::context::get().set_polygon_offset(0.f, 0.f);
'   }
    pso.unbind_vertex_buffers geometryVBO
    gfx_api_context_unbind_index_buffer geometryIndexVBO
End Sub

' lib/ivis_opengl/pielight_convert.h: pielightToRGBAVec4
Private Function pielightToRGBAVec4(ByVal color As Long) As Vector4f
    pielightToRGBAVec4 = glm_vec4((color And &HFF&) / 255!, ((color And &HFF00&) \ &H100&) / 255!, ((color And &HFF0000) \ &H10000) / 255!, (((color And &HFF000000) \ &H1000000) And &HFF&) / 255!)
End Function

Public Function getFogColorVec4() As Vector4f
    Dim fogEnabled As Boolean, fogBegin As Single, fogEnd As Single, fogColour As Long
    renderState_get fogEnabled, fogBegin, fogEnd, fogColour
    getFogColorVec4 = pielightToRGBAVec4(fogColour)
End Function

Private Sub drawTerrainCombinedmpl(ByVal psoType As pso_type, ModelViewProjection As mat4, ViewMatrix As mat4, ModelUVLightmap As mat4, cameraPos As Vector3f, sunPos As Vector3f, shadowCascades As ShadowCascadesInfo)
    Dim fogEnabled As Boolean, fogBegin As Single, fogEnd As Single, fogColour As Long
    Dim pso As pipeline_state_helper, groundScale As mat4, i As Long, w As Long, h As Long
    Dim size As Long, offset As Long, x As Long, y As Long, drawSector As Boolean
    renderState_get fogEnabled, fogBegin, fogEnd, fogColour
    Set pso = gfx_api_pso_get(psoType)
    pso.bind
    pso.bind_textures _
        lightmap_texture, _
        groundTexArr, groundNormalArr, groundSpecularArr, groundHeightArr, _
        decalTexArr, decalNormalArr, decalSpecularArr, decalHeightArr, _
        gfx_api_context_getDepthTexture()
    pso.bind_vertex_buffers terrainDecalVBO
    groundScale = glm_mat4(0!)
    For i = 0 To getNumGroundTypes() - 1
        groundScale.m(i) = 1! / (getGroundType(i).textureSize * world_coord(1)) ' groundScale[i/4][i%4]
    Next

    gfx_api_context_getDrawableDimensions w, h
    tcUniforms.ModelViewProjectionMatrix = ModelViewProjection
    tcUniforms.ViewMatrix = ViewMatrix
    tcUniforms.ModelUVLightmapMatrix = ModelUVLightmap
    For i = 0 To WZ_MAX_SHADOW_CASCADES - 1
        tcUniforms.ShadowMapMVPMatrix(i) = shadowCascades.shadowMVPMatrix(i)
    Next
    tcUniforms.groundScale = groundScale
    tcUniforms.cameraPos = glm_vec4_v3(cameraPos, 0!)
    tcUniforms.sunPos = glm_vec4_v3(glm_normalize3(sunPos), 0!)
    pie_GetLighting0 LIGHT_EMISSIVE, VarPtr(tcUniforms.emissiveLight)
    pie_GetLighting0 LIGHT_AMBIENT, VarPtr(tcUniforms.ambientLight)
    pie_GetLighting0 LIGHT_DIFFUSE, VarPtr(tcUniforms.diffuseLight)
    pie_GetLighting0 LIGHT_SPECULAR, VarPtr(tcUniforms.specularLight)
    tcUniforms.fog_colour = getFogColorVec4()
    tcUniforms.ShadowMapCascadeSplits = glm_vec4(shadowCascades.shadowCascadeSplit(0), shadowCascades.shadowCascadeSplit(1), shadowCascades.shadowCascadeSplit(2), pie_getPerspectiveZFar())
    tcUniforms.ShadowMapSize = shadowCascades.shadowMapSize
    tcUniforms.fog_enabled = -CLng(fogEnabled)
    tcUniforms.fog_begin = fogBegin
    tcUniforms.fog_end = fogEnd
    tcUniforms.quality = terrainShaderQuality
    tcUniforms.viewportWidth = w
    tcUniforms.viewportheight = h
    tcUniforms.unused2 = 0!
    ' getCurrentLightingManager().getPointLightBuckets(): its arrays straight into the uniforms
    tcUniforms.bucketDimensionUsed = ILightingManager_getPointLightBuckets(getCurrentLightingManager(), VarPtr(tcUniforms.PointLightsPosition(0)), VarPtr(tcUniforms.PointLightsColorAndEnergy(0)), VarPtr(tcUniforms.bucketOffsetAndSize(0)), VarPtr(tcUniforms.indexed_lights(0)))
    pso.set_uniforms VarPtr(tcUniforms), LenB(tcUniforms)

    size = 0
    offset = 0
    For x = 0 To xSectors - 1
        For y = 0 To ySectors
            drawSector = False
            If y < ySectors Then drawSector = sectors(x * ySectors + y).draw
            If drawSector Then
                If offset + size = sectors(x * ySectors + y).terrainAndDecalOffset Then
                    ' append
                    size = size + sectors(x * ySectors + y).terrainAndDecalSize
                    GoTo continue_
                End If
            End If
            ' can't append, so draw what we have and start anew
            If size > 0 Then
                pso.draw size, offset
            End If
            size = 0
            If drawSector Then
                offset = sectors(x * ySectors + y).terrainAndDecalOffset
                size = sectors(x * ySectors + y).terrainAndDecalSize
            End If
continue_:
        Next
    Next
    pso.unbind_vertex_buffers terrainDecalVBO
End Sub

Private Sub drawTerrainCombined(ModelViewProjection As mat4, ViewMatrix As mat4, ModelUVLightmap As mat4, cameraPos As Vector3f, sunPos As Vector3f, shadowCascades As ShadowCascadesInfo)
    Select Case terrainShaderQuality
    Case CLASSIC
        drawTerrainCombinedmpl TerrainCombined_Classic, ModelViewProjection, ViewMatrix, ModelUVLightmap, cameraPos, sunPos, shadowCascades
    Case MEDIUM
        drawTerrainCombinedmpl TerrainCombined_Medium, ModelViewProjection, ViewMatrix, ModelUVLightmap, cameraPos, sunPos, shadowCascades
    Case NORMAL_MAPPING
        drawTerrainCombinedmpl TerrainCombined_High, ModelViewProjection, ViewMatrix, ModelUVLightmap, cameraPos, sunPos, shadowCascades
    Case UNINITIALIZED_PICK_DEFAULT
        ' should not happen
    End Select
End Sub

' (VB) lightMap is getCurrentLightmapData()
Public Sub perFrameTerrainUpdates()
    '/////////////////////////////////
    ' set up the lightmap texture

    ' we limit the framerate of the lightmap, because updating a texture is an expensive operation
    If Not ULT(realTime - lightmapLastUpdate, LIGHTMAP_REFRESH) Then
        lightmapLastUpdate = realTime
        updateLightMap

        lightmap_texture.upload 0, lightmapPixmap
    End If

    '/////////////////////////////////
    ' terrain culling
    cullTerrain
End Sub

Public Function getTerrainLightmapTexture() As gfx_api_texture
    Set getTerrainLightmapTexture = lightmap_texture
End Function

Public Function getModelUVLightmapMatrix() As mat4
    getModelUVLightmapMatrix = lightmapValues.ModelUVLightmap
End Function

Public Sub drawTerrainDepthOnly(mvp As mat4)
    drawDepthOnlyForDepthMap mvp, lightmapValues.paramsXLight, lightmapValues.paramsYLight, False
End Sub

' Update the lightmap and draw the terrain and decals.
' This function first draws the terrain in black, and then uses additive blending to put the terrain layers
' on it one by one. Finally the decals are drawn.
Public Sub drawTerrain(mvp As mat4, viewMatrix As mat4, cameraPos As Vector3f, sunPos As Vector3f, shadowCascades As ShadowCascadesInfo)
    If True Then
        '////////////////////////////////////
        ' canvas to draw on
        drawDepthOnly mvp, lightmapValues.paramsXLight, lightmapValues.paramsYLight, True
    End If

    If terrainShaderQuality = CLASSIC Then
        ' draw water *first*
        drawWaterClassic mvp, lightmapValues.ModelUVLightmap, cameraPos, sunPos
    End If

    '/////////////////////////////////
    ' terrain + decals
    drawTerrainCombined mvp, viewMatrix, lightmapValues.ModelUVLightmap, cameraPos, sunPos, shadowCascades
End Sub

' (VB) glm::transpose(glm::mat4(paramsX, paramsY, glm::vec4(0,0,1,0), glm::vec4(0,0,0,1)))
Private Function waterUVMatrix(paramsX As Vector4f, paramsY As Vector4f) As mat4
    waterUVMatrix = glm_transpose(glm_mat4_cols(paramsX, paramsY, glm_vec4(0!, 0!, 1!, 0!), glm_vec4(0!, 0!, 0!, 1!)))
End Function

' Draw the water.
' sunPos and cameraPos in Model=WorldSpace
Private Sub drawWaterNormalImpl(ByVal psoType As pso_type, ModelViewProjection As mat4, cameraPos As Vector3f, sunPos As Vector3f)
    Dim paramsX As Vector4f, paramsY As Vector4f, paramsX2 As Vector4f, paramsY2 As Vector4f
    Dim ModelUV1 As mat4, ModelUV2 As mat4, fogEnabled As Boolean, fogBegin As Single, fogEnd As Single, fogColour As Long
    Dim cb As constant_buffer_type_SHADER_WATER, pso As pipeline_state_helper, x As Long, y As Long, s As Long
    If waterIndexVBO Is Nothing Then
        Exit Sub ' no water
    End If

    paramsX = glm_vec4(0!, 0!, -1! / world_coord(4), 0!)
    paramsY = glm_vec4(1! / world_coord(4), 0!, 0!, 0!)
    paramsX2 = glm_vec4(0!, 0!, -1! / world_coord(5), 0!)
    paramsY2 = glm_vec4(1! / world_coord(5), 0!, 0!, 0!)
    ModelUV1 = glm_mat4_mul(glm_translate(glm_vec3(waterOffset, 0!, 0!)), waterUVMatrix(paramsX, paramsY))
    ModelUV2 = waterUVMatrix(paramsX2, paramsY2)
    renderState_get fogEnabled, fogBegin, fogEnd, fogColour

    If Not (Not waterTexturesNormal.tex1 Is Nothing And Not waterTexturesNormal.tex2 Is Nothing) Then ASSERT_FAILURE "waterTexturesNormal.tex1 && waterTexturesNormal.tex2", "src/terrain.cpp:1574", "drawWaterNormalImpl", "Failed to load water texture": Exit Sub
    Set pso = gfx_api_pso_get(psoType)
    pso.bind

    pso.bind_textures _
        waterTexturesNormal.tex1, waterTexturesNormal.tex2, _
        lightmap_texture
    pso.bind_vertex_buffers waterVBO
    cb.ModelViewProjectionMatrix = ModelViewProjection
    cb.ModelUVLightmapMatrix = lightmapValues.ModelUVLightmap
    cb.ModelUV1Matrix = ModelUV1
    cb.ModelUV2Matrix = ModelUV2
    cb.cameraPos = glm_vec4_v3(cameraPos, 0!)
    cb.sunPos = glm_vec4_v3(glm_normalize3(sunPos), 0!)
    pie_GetLighting0 LIGHT_EMISSIVE, VarPtr(cb.emissiveLight)
    pie_GetLighting0 LIGHT_AMBIENT, VarPtr(cb.ambientLight)
    pie_GetLighting0 LIGHT_DIFFUSE, VarPtr(cb.diffuseLight)
    pie_GetLighting0 LIGHT_SPECULAR, VarPtr(cb.specularLight)
    cb.fog_colour = getFogColorVec4()
    cb.fog_enabled = -CLng(fogEnabled)
    cb.fog_begin = fogBegin
    cb.fog_end = fogEnd
    cb.timeSec = waterOffset * 10!
    pso.bind_constants VarPtr(cb), LenB(cb)

    gfx_api_context_bind_index_buffer waterIndexVBO, index_type_u32

    For x = 0 To xSectors - 1
        For y = 0 To ySectors - 1
            s = x * ySectors + y
            If sectors(s).draw Then
                addDrawRangeElements psoType, _
                    sectors(s).geometryOffset, _
                    sectors(s).geometryOffset + sectors(s).geometrySize, _
                    sectors(s).waterIndexSize, _
                    sectors(s).waterIndexOffset
            End If
        Next
    Next
    finishDrawRangeElements psoType
    pso.unbind_vertex_buffers waterVBO
    gfx_api_context_unbind_index_buffer waterIndexVBO

    ' move the water
    If Not gamePaused() Then
        waterOffset = waterOffset + graphicsTimeAdjustedIncrement(0.1!)
    End If
End Sub

Private Sub drawWaterHighImpl(ByVal psoType As pso_type, ModelViewProjection As mat4, viewMatrix As mat4, cameraPos As Vector3f, sunPos As Vector3f, shadowCascades As ShadowCascadesInfo)
    Dim fogEnabled As Boolean, fogBegin As Single, fogEnd As Single, fogColour As Long
    Dim pso As pipeline_state_helper, x As Long, y As Long, s As Long, i As Long
    If waterIndexVBO Is Nothing Then
        Exit Sub ' no water
    End If

    renderState_get fogEnabled, fogBegin, fogEnd, fogColour

    If Not (Not waterTexturesHigh.tex Is Nothing) Then ASSERT_FAILURE "waterTexturesHigh.tex", "src/terrain.cpp:1630", "drawWaterHighImpl", "Failed to load water textures": Exit Sub
    Set pso = gfx_api_pso_get(psoType)
    pso.bind

    pso.bind_textures _
        waterTexturesHigh.tex, _
        waterTexturesHigh.tex_nm, _
        waterTexturesHigh.tex_sm, _
        lightmap_texture, _
        gfx_api_context_getDepthTexture()
    pso.bind_vertex_buffers waterVBO
    whUniforms.ModelViewProjectionMatrix = ModelViewProjection
    whUniforms.ViewMatrix = viewMatrix
    whUniforms.ModelUVLightmapMatrix = lightmapValues.ModelUVLightmap
    For i = 0 To WZ_MAX_SHADOW_CASCADES - 1
        whUniforms.ShadowMapMVPMatrix(i) = shadowCascades.shadowMVPMatrix(i)
    Next
    whUniforms.cameraPos = glm_vec4_v3(cameraPos, 0!)
    whUniforms.sunPos = glm_vec4_v3(glm_normalize3(sunPos), 0!)
    pie_GetLighting0 LIGHT_EMISSIVE, VarPtr(whUniforms.emissiveLight)
    pie_GetLighting0 LIGHT_AMBIENT, VarPtr(whUniforms.ambientLight)
    pie_GetLighting0 LIGHT_DIFFUSE, VarPtr(whUniforms.diffuseLight)
    pie_GetLighting0 LIGHT_SPECULAR, VarPtr(whUniforms.specularLight)
    whUniforms.fog_colour = getFogColorVec4()
    whUniforms.ShadowMapCascadeSplits = glm_vec4(shadowCascades.shadowCascadeSplit(0), shadowCascades.shadowCascadeSplit(1), shadowCascades.shadowCascadeSplit(2), pie_getPerspectiveZFar())
    whUniforms.ShadowMapSize = shadowCascades.shadowMapSize
    whUniforms.fog_enabled = -CLng(fogEnabled)
    whUniforms.fog_begin = fogBegin
    whUniforms.fog_end = fogEnd
    whUniforms.timeSec = waterOffset * 10!
    pso.bind_constants VarPtr(whUniforms), LenB(whUniforms)

    gfx_api_context_bind_index_buffer waterIndexVBO, index_type_u32

    For x = 0 To xSectors - 1
        For y = 0 To ySectors - 1
            s = x * ySectors + y
            If sectors(s).draw Then
                addDrawRangeElements psoType, _
                    sectors(s).geometryOffset, _
                    sectors(s).geometryOffset + sectors(s).geometrySize, _
                    sectors(s).waterIndexSize, _
                    sectors(s).waterIndexOffset
            End If
        Next
    Next
    finishDrawRangeElements psoType
    pso.unbind_vertex_buffers waterVBO
    gfx_api_context_unbind_index_buffer waterIndexVBO

    ' move the water
    If Not gamePaused() Then
        waterOffset = waterOffset + graphicsTimeAdjustedIncrement(0.1!)
    End If
End Sub

Public Sub drawWaterClassic(ModelViewProjection As mat4, ModelUVLightmap As mat4, cameraPos As Vector3f, sunPos As Vector3f)
    Dim paramsX As Vector4f, paramsY As Vector4f, paramsX2 As Vector4f, paramsY2 As Vector4f
    Dim ModelUV1 As mat4, ModelUV2 As mat4, fogEnabled As Boolean, fogBegin As Single, fogEnd As Single, fogColour As Long
    Dim cb As constant_buffer_type_SHADER_WATER_CLASSIC, pso As pipeline_state_helper, pWaterTexture As gfx_api_texture
    Dim x As Long, y As Long, s As Long
    If waterIndexVBO Is Nothing Then
        Exit Sub ' no water
    End If

    paramsX = glm_vec4(0!, 0!, -1! / world_coord(4), 0!)
    paramsY = glm_vec4(1! / world_coord(4), 0!, 0!, 0!)
    paramsX2 = glm_vec4(0!, 0!, -1! / world_coord(5), 0!)
    paramsY2 = glm_vec4(1! / world_coord(5), 0!, 0!, 0!)
    ModelUV1 = glm_mat4_mul(glm_translate(glm_vec3(waterOffset, 0!, 0!)), waterUVMatrix(paramsX, paramsY))
    ModelUV2 = waterUVMatrix(paramsX2, paramsY2)
    renderState_get fogEnabled, fogBegin, fogEnd, fogColour

    Set pWaterTexture = getWaterClassicTexture()
    If Not (Not pWaterTexture Is Nothing) Then ASSERT_FAILURE "pWaterTexture != nullptr", "src/terrain.cpp:1691", "drawWaterClassic", "Failed to load water legacy texture?"

    Set pso = gfx_api_pso_get(WaterClassicPSO)
    pso.bind

    pso.bind_textures lightmap_texture, pWaterTexture
    pso.bind_vertex_buffers waterVBO
    cb.ModelViewProjectionMatrix = ModelViewProjection
    cb.ModelUVLightmapMatrix = ModelUVLightmap
    cb.ShadowMapMVPMatrix = glm_mat4(1!) ' TODO
    cb.ModelUV1Matrix = ModelUV1
    cb.ModelUV2Matrix = ModelUV2
    cb.cameraPos = glm_vec4_v3(cameraPos, 0!)
    cb.sunPos = glm_vec4_v3(glm_normalize3(sunPos), 0!)
    cb.fog_colour = getFogColorVec4()
    cb.fog_enabled = -CLng(fogEnabled)
    cb.fog_begin = fogBegin
    cb.fog_end = fogEnd
    cb.timeSec = waterOffset * 10!
    pso.set_uniforms VarPtr(cb), LenB(cb)

    gfx_api_context_bind_index_buffer waterIndexVBO, index_type_u32

    For x = 0 To xSectors - 1
        For y = 0 To ySectors - 1
            s = x * ySectors + y
            If sectors(s).draw Then
                addDrawRangeElements WaterClassicPSO, _
                    sectors(s).geometryOffset, _
                    sectors(s).geometryOffset + sectors(s).geometrySize, _
                    sectors(s).waterIndexSize, _
                    sectors(s).waterIndexOffset
            End If
        Next
    Next
    finishDrawRangeElements WaterClassicPSO
    pso.unbind_vertex_buffers waterVBO
    gfx_api_context_unbind_index_buffer waterIndexVBO

    ' move the water
    If Not gamePaused() Then
        waterOffset = waterOffset + graphicsTimeAdjustedIncrement(0.1!)
    End If
End Sub

Public Sub drawWater(ModelViewProjection As mat4, viewMatrix As mat4, cameraPos As Vector3f, sunPos As Vector3f, shadowCascades As ShadowCascadesInfo)
    Select Case terrainShaderQuality
    Case CLASSIC
        ' already drawn
        Exit Sub
    Case MEDIUM
        drawWaterNormalImpl WaterPSO, ModelViewProjection, cameraPos, sunPos
        Exit Sub
    Case NORMAL_MAPPING
        drawWaterHighImpl WaterHighPSO, ModelViewProjection, viewMatrix, cameraPos, sunPos, shadowCascades
        Exit Sub
    Case UNINITIALIZED_PICK_DEFAULT
        ' should not happen
    End Select
End Sub

' MARK: Terrain shader quality / appearance

Public Function determineTerrainShaderRequirementsMet() As Boolean
    Dim maxArrayTextureLayers As Long
    If Not gfx_api_context_supports2DTextureArrays() Then
        debug_ 1758, LOG_ERROR, "determineTerrainShaderRequirementsMet", "Graphics context: supports2DTextureArrays = false"
        determineTerrainShaderRequirementsMet = False
        Exit Function
    End If
    If Not gfx_api_context_supportsIntVertexAttributes() Then
        debug_ 1763, LOG_ERROR, "determineTerrainShaderRequirementsMet", "Graphics context: supportsIntVertexAttributes = false"
        determineTerrainShaderRequirementsMet = False
        Exit Function
    End If
    maxArrayTextureLayers = gfx_api_context_get_context_value(context_value_MAX_ARRAY_TEXTURE_LAYERS)
    If maxArrayTextureLayers < 256 Then ' "big enough" value
        debug_ 1769, LOG_ERROR, "determineTerrainShaderRequirementsMet", "Graphics context: MAX_ARRAY_TEXTURE_LAYERS is insufficient: %d", maxArrayTextureLayers
        determineTerrainShaderRequirementsMet = False
        Exit Function
    End If

    determineTerrainShaderRequirementsMet = True
End Function

Public Function getTerrainShaderQuality() As Long
    getTerrainShaderQuality = terrainShaderQuality
End Function

' bool setTerrainShaderQuality(TerrainShaderQuality newValue, bool force, bool forceReloadTextures = false)
Public Function setTerrainShaderQuality_force(ByVal newValue As Long, ByVal force As Boolean, Optional ByVal forceReloadTextures As Boolean = False) As Boolean
    Dim success As Boolean, priorValue As Long, supportedModes() As Long, supportedModes_size As Long
    Dim debugQualityUint As Long, terrainQualityUintStr As String, priorNumGroundTypes As Long
    If Not initializedTerrainShaderType Then
        ' cache the desired value for when terrain shader type is initialized (post graphics context initialization)
        terrainShaderQuality = newValue
        setTerrainShaderQuality_force = True
        Exit Function
    End If

    If Not force And newValue = terrainShaderQuality Then
        setTerrainShaderQuality_force = True
        Exit Function
    End If

    success = False
    priorValue = terrainShaderQuality

    ' Check whether the new shader can be used
    ' - all quality modes are supported by shader
    ' - but check if texture pack is available
    If isSupportedTerrainShaderQualityOption(newValue) Then
        terrainShaderQuality = newValue
        success = True
    Else
        ' texture pack is not available
        terrainShaderQuality = priorValue
        If newValue <> priorValue Then
            debug_ 1812, LOG_INFO, "setTerrainShaderQuality", "Missing texture pack for %s terrain quality - reverting to %s terrain quality", to_display_string(newValue), to_display_string(priorValue)
            success = False
        Else
            ' need to pick another supported value
            supportedModes_size = getAvailableTerrainShaderQualityTextures(supportedModes)
            If supportedModes_size > 0 Then
                debug_ 1821, LOG_INFO, "setTerrainShaderQuality", "Missing texture pack for %s terrain quality - reverting to %s terrain quality", to_display_string(newValue), to_display_string(supportedModes(supportedModes_size - 1))
                terrainShaderQuality = supportedModes(supportedModes_size - 1)
                success = True
            End If
        End If
        newValue = terrainShaderQuality
    End If

    If success Then
        debugQualityUint = terrainShaderQuality
        terrainQualityUintStr = CStr(debugQualityUint)
        crashHandlingProviderSetTag "wz.terrain_quality", terrainQualityUintStr

        dirtyAllSectors

        ' re-load terrain textures
        If Not rebuildExistingSearchPathWithGraphicsOptionChange() Then
            debug_ 1840, LOG_INFO, "setTerrainShaderQuality", "Failed to swap terrain texture overrides"
        End If
        ' always re-load the tile textures (these change between all modes)
        priorNumGroundTypes = getNumGroundTypes()
        reloadTileTextures
        mapReloadGroundTypes
        If priorValue <> terrainShaderQuality _
           Or priorNumGroundTypes <> getNumGroundTypes() _
           Or forceReloadTextures Then
            reloadTerrainTextures
        End If
    End If

    setTerrainShaderQuality_force = success
End Function

Public Function setTerrainShaderQuality(ByVal newValue As Long) As Boolean
    setTerrainShaderQuality = setTerrainShaderQuality_force(newValue, False)
End Function

' std::vector<TerrainShaderQuality>: the values in result (0-based), their number returned
Public Function getAllTerrainShaderQualityOptions(result() As Long) As Long
    ReDim result(0 To 2)
    result(0) = CLASSIC
    result(1) = MEDIUM
    result(2) = NORMAL_MAPPING
    getAllTerrainShaderQualityOptions = 3
End Function

Public Function isSupportedTerrainShaderQualityOption(ByVal value As Long) As Boolean
    Dim supportedModes() As Long, supportedModes_size As Long, i As Long
    If Not initializedTerrainShaderType Then ASSERT_FAILURE "initializedTerrainShaderType", "src/terrain.cpp:1869", "isSupportedTerrainShaderQualityOption", "Should only be called after graphics context initialization": isSupportedTerrainShaderQualityOption = False: Exit Function
    supportedModes_size = getAvailableTerrainShaderQualityTextures(supportedModes)
    ' all quality modes are supported by shader - check if textures are available
    For i = 0 To supportedModes_size - 1
        If supportedModes(i) = value Then
            isSupportedTerrainShaderQualityOption = True
            Exit Function
        End If
    Next
End Function

' (VB) std::string::rfind(prefix, 0) == 0
Private Function startsWith(ByVal s As String, ByVal prefix As String) As Boolean
    startsWith = (Left$(s, Len(prefix)) = prefix)
End Function

' (VB) the build is 32-bit (INTPTR_MAX <= INT32_MAX): the #if branches of a 32-bit build
Private Function determineDefaultTerrainQuality() As Long
    Dim systemRAMinMiB As Double, backendInfo As Object, estimatedVRAMinMiB As Double
    ' Based on system properties, determine a reasonable default (for performance reasons)
    ' (Uses a heuristic based on system RAM, graphics renderer, and estimated VRAM)

    ' If <= 8 GiB system RAM, default to medium ("normal")
    systemRAMinMiB = wzGetCurrentSystemRAM()
    If systemRAMinMiB <= 8192 Then
        debug_ 1884, LOG_INFO, "determineDefaultTerrainQuality", "Due to system RAM (%u MiB), defaulting to terrain quality: Normal", systemRAMinMiB
        determineDefaultTerrainQuality = MEDIUM
        Exit Function
    End If

    ' For specific older integrated cards on OpenGL, default to medium ("normal")
    Set backendInfo = gfx_api_context_getBackendGameInfo()
    If Not backendInfo Is Nothing Then
        If backendInfo.Exists("openGL_renderer") Then
            ' If opengl_renderer starts with "Intel(R) HD Graphics"
            If startsWith(backendInfo.Item("openGL_renderer"), "Intel(R) HD Graphics") Then
                ' default to medium ("normal")
                determineDefaultTerrainQuality = MEDIUM
                Exit Function
            End If
            ' If opengl_renderer starts with "Intel(R) Graphics Media Accelerator"
            If startsWith(backendInfo.Item("openGL_renderer"), "Intel(R) Graphics Media Accelerator") Then
                ' default to medium ("normal")
                determineDefaultTerrainQuality = MEDIUM
                Exit Function
            End If
        End If

        ' If it's the null backend (headless mode)
        If backendInfo.Exists("null_gfx_backend") Then
            ' default to medium ("normal")
            debug_ 1912, LOG_INFO, "determineDefaultTerrainQuality", "Due to null gfx backend, defaulting to terrain quality: Normal"
            determineDefaultTerrainQuality = MEDIUM
            Exit Function
        End If
    End If

    ' Try to get the estimated available VRAM
    estimatedVRAMinMiB = gfx_api_context_get_estimated_vram_mb(False)
    If estimatedVRAMinMiB > 0 Then
        ' If estimated VRAM < 2 GiB
        If estimatedVRAMinMiB < 2048 Then
            ' default to medium ("normal")
            debug_ 1924, LOG_INFO, "determineDefaultTerrainQuality", "Due to estimated VRAM (%u MiB), defaulting to terrain quality: Normal", estimatedVRAMinMiB
            determineDefaultTerrainQuality = MEDIUM
            Exit Function
        End If
    End If

    ' On 32-bit builds, default to MEDIUM
    debug_ 1931, LOG_INFO, "determineDefaultTerrainQuality", "32-bit build defaulting to terrain quality: Normal"
    determineDefaultTerrainQuality = MEDIUM
End Function

Private Function determineDefaultTerrainMappingTextureSize() As Long
    Dim systemRAMinMiB As Double, backendInfo As Object, estimatedVRAMinMiB As Double
    If determineDefaultTerrainQuality() = MEDIUM Then
        ' system checks yield medium terrain quality as the default
        ' so default the terrain mapping texture size to 512 (in case the user switches to NORMAL_MAPPING)
        determineDefaultTerrainMappingTextureSize = 512
        Exit Function
    End If

    ' If < 8 GiB system RAM, default to 512
    systemRAMinMiB = wzGetCurrentSystemRAM()
    If systemRAMinMiB < 8192 Then
        debug_ 1952, LOG_INFO, "determineDefaultTerrainMappingTextureSize", "Due to system RAM (%u MiB), defaulting to terrain mapping texture size: 512", systemRAMinMiB
        determineDefaultTerrainMappingTextureSize = 512
        Exit Function
    End If

    ' For specific older integrated cards on OpenGL, default to 512
    Set backendInfo = gfx_api_context_getBackendGameInfo()
    If Not backendInfo Is Nothing Then
        If backendInfo.Exists("openGL_renderer") Then
            ' If opengl_renderer starts with "Intel(R) HD Graphics"
            If startsWith(backendInfo.Item("openGL_renderer"), "Intel(R) HD Graphics") Then
                determineDefaultTerrainMappingTextureSize = 512
                Exit Function
            End If
            ' If opengl_renderer starts with "Intel(R) Graphics Media Accelerator"
            If startsWith(backendInfo.Item("openGL_renderer"), "Intel(R) Graphics Media Accelerator") Then
                determineDefaultTerrainMappingTextureSize = 512
                Exit Function
            End If
        End If
    End If

    ' Get the estimated *dedicated* VRAM
    estimatedVRAMinMiB = gfx_api_context_get_estimated_vram_mb(True)
    If estimatedVRAMinMiB <= 4096 Then
        ' If estimated dedicated VRAM value is <= 4 GiB (or cannot be determined), default to 512
        debug_ 1978, LOG_INFO, "determineDefaultTerrainMappingTextureSize", "Due to estimated dedicated VRAM (%u MiB), defaulting to terrain mapping texture size: 512", estimatedVRAMinMiB
        determineDefaultTerrainMappingTextureSize = 512
        Exit Function
    End If

    ' On 32-bit builds, default to 512
    debug_ 1984, LOG_INFO, "determineDefaultTerrainMappingTextureSize", "32-bit build defaulting to terrain mapping texture size: 512"
    determineDefaultTerrainMappingTextureSize = 512
End Function

Public Sub initTerrainShaderType()
    If Not determineTerrainShaderRequirementsMet() Then
        crashHandlingProviderCaptureException "initTerrainShaderType", "NewTerrainRendererUnsupported", "", False, True
        debug_ 1997, LOG_FATAL, "initTerrainShaderType", "Your system does not support the terrain renderer. Please check your graphics drivers."
        exit_ 3                         ' abort()
    End If

    initializedTerrainShaderType = True
    If terrainShaderQuality = UNINITIALIZED_PICK_DEFAULT Then
        terrainShaderQuality = determineDefaultTerrainQuality()
        debug_ 2005, LOG_INFO, "initTerrainShaderType", "Defaulting terrain quality to: %s", to_display_string(terrainShaderQuality)
    End If
    If maxTerrainMappingTextureSize = 0 Then
        maxTerrainMappingTextureSize = determineDefaultTerrainMappingTextureSize()
    End If
    setTerrainShaderQuality_force terrainShaderQuality, True ' checks and resets unsupported values
End Sub

Public Function to_display_string(ByVal value As Long) As String
    Select Case value
    Case UNINITIALIZED_PICK_DEFAULT
        to_display_string = ""
    Case CLASSIC
        to_display_string = "Classic"
    Case MEDIUM
        to_display_string = "Normal"
    Case NORMAL_MAPPING
        to_display_string = "Remastered (HQ)"
    Case Else
        to_display_string = "" ' silence warning
    End Select
End Function
