Attribute VB_Name = "lighting_c"
' src/lighting.cpp, lighting.h: the sun, the tiles' illumination and ambient
' occlusion (display only), the fog's distance and colour, the droids'
' illumination. rendering1999::LightingManager is the class
' rendering1999_LightingManager (it implements WzGfx's ILightingManager).
' pienormalize.h's normalise and pie_SurfaceNormal3fv are this module's own
' copies (a header's inlines: CONVENTIONS). The display's float arithmetic is
' plain Single arithmetic, but calcDroidIllumination's, which the simulation's
' tests see (the droids' illumination), goes through float_vb as C++ rounds it.
Option Explicit

' These magic values determine the fog
Private Const FOG_ALTITUDE_COEFFICIENT As Single = 1.3!

Private Const MIN_DROID_LIGHT_LEVEL As Long = 96
Private Const DROID_SEEK_LIGHT_SPEED As Long = 2

' The vector that holds the sun's lighting direction - planar
Private theSun As Vector3f                      ' (0.f, 0.f, 0.f)
Private theSun_ForTileIllumination As Vector3f  ' (0.f, 0.f, 0.f)

' pienormalize.h: glm::dot(a, a) > 0.f ? glm::normalize(a) : a
Private Function normalise(a As Vector3f) As Vector3f
    If glm_dot3(a, a) > 0! Then
        normalise = glm_normalize3(a)
    Else
        normalise = a
    End If
End Function

' pienormalize.h: Calculate surface normal
Private Function pie_SurfaceNormal3fv(p1 As Vector3f, p2 As Vector3f, p3 As Vector3f) As Vector3f
    Dim a As Vector3f, b As Vector3f, c As Vector3f
    a = glm_vec3_sub(p3, p1)
    b = glm_vec3_sub(p2, p1)
    c = glm_cross(a, b)
    pie_SurfaceNormal3fv = normalise(c)
End Function

Public Sub setTheSun(newSun As Vector3f)
    Dim oldSun As Vector3f, n As Vector3f
    oldSun = theSun
    n = normalise(newSun)
    theSun = glm_vec3_mul(n, CSng(FP12_MULTIPLIER))
    theSun_ForTileIllumination = glm_vec3(-theSun.x, -theSun.y, theSun.z)
    If oldSun.x <> theSun.x Or oldSun.y <> theSun.y Or oldSun.z <> theSun.z Then
        ' The sun has changed - must relcalulate lighting
        initLighting 0, 0, mapWidth, mapHeight
    End If
End Sub

Public Function getTheSun() As Vector3f
    getTheSun = theSun
End Function

'****************************************************************************
'
' SOURCE
'
'****************************************************************************

'By passing in params - it means that if the scroll limits are changed mid-mission
'we can re-do over the area that hasn't been seen
Public Sub initLighting(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    Dim i As Long, j As Long, psTile As Long
    ' quick check not trying to go off the map - don't need to check for < 0 since UWORD's!!
    If x1 > mapWidth Or x2 > mapWidth Or y1 > mapHeight Or y2 > mapHeight Then
        ASSERT_FAILURE "false", "src/lighting.cpp:88", "initLighting", "initLighting: coords off edge of map"
        Exit Sub
    End If

    For j = y1 To y2 - 1
        For i = x1 To x2 - 1
            psTile = mapTile(i, j)

            ' always make the edge tiles dark
            If i = 0 Or j = 0 Or i >= mapWidth - 1 Or j >= mapHeight - 1 Then
                psMapTiles(psTile).illumination = 16
                psMapTiles(psTile).ambientOcclusion = 16
            Else
                calcTileIllum i, j
            End If
            ' Basically darkens down the tiles that are outside the scroll
            ' limits - thereby emphasising the cannot-go-there-ness of them
            If i < scrollMinX + 4 Or i > scrollMaxX - 4 _
               Or j < scrollMinY + 4 Or j > scrollMaxY - 4 Then
                psMapTiles(psTile).illumination = psMapTiles(psTile).illumination \ 3
                psMapTiles(psTile).ambientOcclusion = psMapTiles(psTile).ambientOcclusion \ 3
            End If
        Next
    Next
End Sub

' For display purposes only (*NOT* for use in game state calculations)
Private Sub normalsOnTile(ByVal tileX As Long, ByVal tileY As Long, ByVal quadrant As Long, numNormals As Long, normals() As Vector3f)
    Dim tiles(0 To 1, 0 To 1) As Vector2i, psTiles(0 To 1, 0 To 1) As Long, corners(0 To 1, 0 To 1) As Vector3f
    Dim i As Long, j As Long, flipped As Long

    For j = 0 To 1
        For i = 0 To 1
            tiles(i, j).x = tileX + i
            tiles(i, j).y = tileY + j
            ' Get a pointer to our tile
            ' And to the ones to the east, south and southeast of it
            psTiles(i, j) = mapTile(tiles(i, j).x, tiles(i, j).y)
            corners(i, j) = glm_vec3(world_coord(tiles(i, j).x), world_coord(tiles(i, j).y), psMapTiles(psTiles(i, j)).height)
        Next
    Next

    If TRI_FLIPPED(psTiles(0, 0)) Then flipped = 10 Else flipped = 0

    Select Case quadrant + flipped
    ' Not flipped.
    Case 0, 2
        ' Otherwise, it's not flipped and so two triangles
        normals(numNormals) = pie_SurfaceNormal3fv(corners(0, 0), corners(1, 0), corners(1, 1))
        numNormals = numNormals + 1
        normals(numNormals) = pie_SurfaceNormal3fv(corners(0, 0), corners(1, 1), corners(0, 1))
        numNormals = numNormals + 1
    Case 1
        normals(numNormals) = pie_SurfaceNormal3fv(corners(0, 0), corners(1, 1), corners(0, 1))
        numNormals = numNormals + 1
    Case 3
        normals(numNormals) = pie_SurfaceNormal3fv(corners(0, 0), corners(1, 0), corners(1, 1))
        numNormals = numNormals + 1
    ' Flipped.
    Case 10
        ' Is it flipped? In this case one triangle
        normals(numNormals) = pie_SurfaceNormal3fv(corners(1, 0), corners(1, 1), corners(0, 1))
        numNormals = numNormals + 1
    Case 12
        normals(numNormals) = pie_SurfaceNormal3fv(corners(0, 0), corners(1, 0), corners(0, 1))
        numNormals = numNormals + 1
    Case 11, 13
        ' Is it flipped? In this case two triangles
        normals(numNormals) = pie_SurfaceNormal3fv(corners(0, 0), corners(1, 0), corners(0, 1))
        numNormals = numNormals + 1
        normals(numNormals) = pie_SurfaceNormal3fv(corners(1, 0), corners(1, 1), corners(0, 1))
        numNormals = numNormals + 1
    Case Else
        ASSERT_FAILURE "false", "src/lighting.cpp:169", "normalsOnTile", "Invalid quadrant in lighting code"
    End Select ' end switch
End Sub

' (VB) clip<float>
Private Function clipf(ByVal x As Single, ByVal minVal As Single, ByVal maxVal As Single) As Single
    If x < minVal Then
        clipf = minVal
    ElseIf x > maxVal Then
        clipf = maxVal
    Else
        clipf = x
    End If
End Function

' For display purposes only (*NOT* for use in game state calculations)
Private Sub calcTileIllum(ByVal tileX As Long, ByVal tileY As Long)
    Dim numNormals As Long ' How many normals have we got?
    Dim normals(0 To 7) As Vector3f ' Maximum 8 possible normals
    Dim finalVector As Vector3f, i As Long, dotProduct As Single, n As Vector3f
    Dim ao As Single, cx As Long, cy As Long, maxX As Long, maxY As Long, height As Single
    Dim dx(0 To 7) As Single, dy(0 To 7) As Single, dir_ As Long, dist As Long, maxTangent As Single, tangent As Single
    Dim tile As Long
    Const I_ As Single = 100!                        ' constexpr float I = 100;
    Const H As Single = I_ * 0.707106781186548!      ' sqrt(1/2)
    Const Dirs As Long = 8

    ' Quadrants look like:-
    '
    '              *
    '              *
    '        0     *    1
    '              *
    '              *
    '    **********V**********
    '              *
    '              *
    '        3     *    2
    '              *
    '              *

    ' Do quadrant 0 - tile that's above and left
    normalsOnTile tileX - 1, tileY - 1, 0, numNormals, normals

    ' Do quadrant 1 - tile that's above and right
    normalsOnTile tileX, tileY - 1, 1, numNormals, normals

    ' Do quadrant 2 - tile that's down and right
    normalsOnTile tileX, tileY, 2, numNormals, normals

    ' Do quadrant 3 - tile that's down and left
    normalsOnTile tileX - 1, tileY, 3, numNormals, normals

    ' The number or normals that we got is in numNormals
    finalVector = glm_vec3(0!, 0!, 0!)
    For i = 0 To numNormals - 1
        finalVector = glm_vec3_add(finalVector, normals(i))
    Next

    n = normalise(finalVector)
    dotProduct = glm_dot3(n, theSun_ForTileIllumination) / 16!

    ' Primitive ambient occlusion calculation.
    ao = 0!
    cx = world_coord(tileX)
    cy = world_coord(tileY)
    maxX = world_coord(mapWidth)
    maxY = world_coord(mapHeight)
    height = map_Height(clip_(cx, 0, maxX), clip_(cy, 0, maxY))
    dx(0) = 0!: dx(1) = H: dx(2) = I_: dx(3) = H: dx(4) = 0!: dx(5) = -H: dx(6) = -I_: dx(7) = -H   ' I sin(2pi dir/Dirs)
    dy(0) = I_: dy(1) = H: dy(2) = 0!: dy(3) = -H: dy(4) = -I_: dy(5) = -H: dy(6) = 0!: dy(7) = H  ' I cos(2pi dir/Dirs)

    For dir_ = 0 To Dirs - 1
        maxTangent = 0!
        For dist = 1 To 8
            tangent = (map_Height(clip_(CLng(Fix(cx + dx(dir_) * dist)), 0, maxX), clip_(CLng(Fix(cy + dy(dir_) * dist)), 0, maxY)) - height) / (I_ * dist)
            If tangent > maxTangent Then maxTangent = tangent
        Next
        ' Ambient light in this direction is proportional to the integral from tan(phi) = tangent to tan(phi) = infinity of dphi cos(phi).
        ' Indefinite integral is sin(phi), so definite integral is 1 - sin(atan(tangent)) = 1 - tangent/sqrt(tangent^2 + 1).
        ao = ao + (1! - maxTangent / Sqr(maxTangent * maxTangent + 1!))
    Next
    ao = ao * (1! / Dirs)
    ao = clipf(ao, 0.25!, 1!)

    tile = mapTile(tileX, tileY)
    psMapTiles(tile).illumination = clip_(CLng(Fix(Abs(dotProduct * ao))), 24, 254)
    psMapTiles(tile).ambientOcclusion = CLng(Fix(clipf(254! * ao, 60!, 254!)))
End Sub

' (VB) the bytes of a PIELIGHT (a Long, r in the low byte)
Private Function PIELIGHT_byte(ByVal colour As Long, ByVal idx As Long) As Long
    Select Case idx
    Case 0: PIELIGHT_byte = colour And &HFF&
    Case 1: PIELIGHT_byte = (colour And &HFF00&) \ &H100&
    Case 2: PIELIGHT_byte = (colour And &HFF0000) \ &H10000
    Case Else: PIELIGHT_byte = ((colour And &HFF000000) \ &H1000000) And &HFF&
    End Select
End Function

Private Sub colourTile(ByVal xIndex As Long, ByVal yIndex As Long, ByVal light_colour As Long, ByVal fraction As Double)
    Dim colour As Long, r As Double, g As Double, b As Double
    colour = LightMap_get(xIndex, yIndex)
    r = PIELIGHT_byte(colour, 0) + PIELIGHT_byte(light_colour, 0) * fraction
    g = PIELIGHT_byte(colour, 1) + PIELIGHT_byte(light_colour, 1) * fraction
    b = PIELIGHT_byte(colour, 2) + PIELIGHT_byte(light_colour, 2) * fraction
    If r > 255 Then r = 255
    If g > 255 Then g = 255
    If b > 255 Then b = 255
    LightMap_set xIndex, yIndex, pal_RGBA(CLng(Fix(r)), CLng(Fix(g)), CLng(Fix(b)), PIELIGHT_byte(colour, 3))
End Sub

' rendering1999::LightingManager::ComputeFrameData (rendering1999_LightingManager forwards here)
' (VB) data and lightmap are getCurrentLightingData() and getCurrentLightmapData() (WzGfx's
' LightingData_lights_* and LightMap_*); the matrix is unused
Public Sub rendering1999_LightingManager_ComputeFrameData(ByVal worldViewProjectionMatrix As Long)
    Dim li As Long, n As Long, px As Long, py As Long, pz As Long, range As Long, colour As Long
    Dim tileX As Long, tileY As Long, rangeSkip As Long, startX As Long, endX As Long, startY As Long, endY As Long
    Dim i As Long, j As Long, position As Vector3i, distToCorner As Long, ratio As Double
    n = LightingData_lights_size()
    For li = 0 To n - 1
        LightingData_lights_at li, px, py, pz, range, colour
        ' Firstly - there's no point processing lights that are off the grid
        If clipXY(px, pz) = False Then
            GoTo continue_
        End If

        tileX = px \ TILE_UNITS
        tileY = pz \ TILE_UNITS
        rangeSkip = CLng(Fix(Sqr(CSng(CDbl(range) * range * 2)) / TILE_UNITS + 1))

        ' Rough guess?
        startX = tileX - rangeSkip
        endX = tileX + rangeSkip
        startY = tileY - rangeSkip
        endY = tileY + rangeSkip

        ' Clip to grid limits
        If startX < 0 Then startX = 0
        If endX < 0 Then endX = 0
        If endX > mapWidth - 1 Then endX = mapWidth - 1
        If startX > endX Then startX = endX
        If startY < 0 Then startY = 0
        If endY < 0 Then endY = 0
        If endY > mapHeight - 1 Then endY = mapHeight - 1
        If startY > endY Then startY = endY

        For i = startX To endX
            For j = startY To endY
                position.x = px
                position.y = py
                position.z = pz
                distToCorner = calcDistToTile(i, j, position)

                ' If we're inside the range of the light
                If distToCorner < range Then
                    ' Find how close we are to it
                    ratio = (100# - ((distToCorner * 100) \ range)) / 100#   ' PERCENT(distToCorner, psLight->range)
                    colourTile i, j, colour, ratio
                End If
            Next
        Next
continue_:
    Next
End Sub

Private Function calcDistToTile(ByVal tileX As Long, ByVal tileY As Long, pos As Vector3i) As Long
    Dim x1 As Long, y1 As Long, z1 As Long

    ' The coordinates of the tile corner
    x1 = tileX * TILE_UNITS
    y1 = map_TileHeight(tileX, tileY)
    z1 = tileY * TILE_UNITS

    calcDistToTile = iHypot3(x1 - pos.x, y1 - pos.y, z1 - pos.z)
End Function

' Sets the begin and end distance for the distance fog (mist)
' It should provide maximum visibility and minimum
' "popping" tiles
Public Sub updateFogDistance(ByVal distance As Single)
    pie_UpdateFogDistance war_getFogStart() + (distance - war_GetMapZoom()) * FOG_ALTITUDE_COEFFICIENT, war_getFogEnd() + (distance - war_GetMapZoom()) * FOG_ALTITUDE_COEFFICIENT
End Sub

Public Sub setDefaultFogColour()
    If Not (StrPtr(tilesetDir) <> 0) Then ASSERT_FAILURE "tilesetDir != nullptr", "src/lighting.cpp:325", "setDefaultFogColour", "Uninitialized tilesetDir"
    Select Case currentMapTileset
    Case MAP_TILESET_ARIZONA
        ' Arizona, and default. = b08f5f (or, 0x78684f)
        pie_SetFogColour psPalette(WZCOL_FOG_ARIZONA)
    Case MAP_TILESET_URBAN
        ' Urban = 0x101040 (or, 0xc9920f)
        pie_SetFogColour psPalette(WZCOL_FOG_URBAN)
    Case MAP_TILESET_ROCKIES
        ' Rockies = 0xb6e1ec
        pie_SetFogColour psPalette(WZCOL_FOG_ROCKIE)
    End Select
End Sub

Public Sub calcDroidIllumination(ByVal psDroid As Long)
    Dim lightVal As Long, presVal As Long, retVal As Long, adjust As Single
    Dim tileX As Long, tileY As Long
    tileX = map_coord(aDroid(psDroid).pos.x)
    tileY = map_coord(aDroid(psDroid).pos.y)

    ' Are we at the edge, or even on the map
    If Not tileOnMap(tileX, tileY) Then
        aDroid(psDroid).illumination = UBYTE_MAX
        Exit Sub
    ElseIf tileX <= 1 Or tileX >= mapWidth - 2 Or tileY <= 1 Or tileY >= mapHeight - 2 Then
        lightVal = psMapTiles(mapTile(tileX, tileY)).illumination
        lightVal = lightVal + MIN_DROID_LIGHT_LEVEL
    Else
        lightVal = psMapTiles(mapTile(tileX, tileY)).illumination + _
                   psMapTiles(mapTile(tileX - 1, tileY)).illumination + _
                   psMapTiles(mapTile(tileX, tileY - 1)).illumination + _
                   psMapTiles(mapTile(tileX + 1, tileY)).illumination + _
                   psMapTiles(mapTile(tileX + 1, tileY + 1)).illumination
        lightVal = lightVal \ 5
        lightVal = lightVal + MIN_DROID_LIGHT_LEVEL
    End If

    ' Saturation
    If lightVal > 255 Then
        lightVal = 255
    End If
    presVal = aDroid(psDroid).illumination
    adjust = fsub(CSng(lightVal), CSng(presVal))
    adjust = fmul(adjust, graphicsTimeAdjustedIncrement(DROID_SEEK_LIGHT_SPEED))
    retVal = CLng(Fix(fadd(CSng(presVal), adjust)))
    If retVal > 255 Then
        retVal = 255
    End If
    aDroid(psDroid).illumination = retVal And &HFF&   ' (UBYTE)
End Sub
