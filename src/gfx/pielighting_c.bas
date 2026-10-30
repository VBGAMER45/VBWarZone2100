Attribute VB_Name = "pielighting_c"
' lib/ivis_opengl/pielighting.cpp/.h: the scene's lights, the light map and
' the lighting managers. getCurrentLightingData() and getCurrentLightmapData()
' are function-local statics: here the module's lights() and lightMap; the
' exe reads and writes them through the LightingData_lights_* and LightMap_*
' procedures (GfxApi). ILightingManager is an interface (ILightingManager.cls)
' with its one virtual method; ComputeFrameData's data and lightmap arguments
' are always the current ones, so the interface method takes only the matrix
' (an address). The base class's currentPointLightBuckets are kept here, for
' the current manager (the only one read; display3d makes a new one each
' frame and calls SetFrameStart first). renderingNew::LightingManager is the
' class renderingNew_LightingManager (made by renderingNew_LightingManager_new);
' its ComputeFrameData is renderingNew_LightingManager_ComputeFrameData here.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Public Type LIGHT
    position As Vector3i                ' = Vector3i(0, 0, 0)
    range As Long                       ' UDWORD
    colour As Long                      ' PIELIGHT
End Type

Public Type PointLightBuckets
    positions(0 To gfx_api_max_lights - 1) As Vector4f
    colorAndEnergy(0 To gfx_api_max_lights - 1) As Vector4f

    ' z and y components are used for padding, keep ivec4 !
    bucketOffsetAndSize(0 To gfx_api_bucket_dimension * gfx_api_bucket_dimension - 1) As Vector4i
    ' Unfortunately due to std140 constraint, we pack indexes in glm::ivec4 and unpack them in shader later
    light_index(0 To gfx_api_max_indexed_lights - 1) As Vector4i

    bucketDimensionUsed As Long         ' = gfx_api::bucket_dimension
End Type

Public Type CalculatedPointLight
    position As Vector3f                ' = glm::vec3(0, 0, 0)
    colour As Vector3f
    range As Single
End Type

' getCurrentLightingData(): static LightingData scene; (its std::vector<LIGHT> lights)
Private lights() As LIGHT
Private lights_size As Long

' getCurrentLightmapData(): static LightMap lightmap;
Private lightMapData() As Long          ' std::unique_ptr<PIELIGHT[]> data
Private mapWidth As Long
Private mapHeight As Long

' ILightingManager::currentPointLightBuckets of the current manager
Public currentPointLightBuckets As PointLightBuckets

' static std::unique_ptr<ILightingManager> lightingManager;
Private lightingManager As ILightingManager

'--------------------------------------------------------------------------
' LightMap

' The shift on a world coordinate to get the tile coordinate
Private Const TILE_SHIFT As Long = 7

' PIELIGHT& LightMap::operator()(int32_t x, int32_t y): (VB) the index of the element
Private Function LightMap_index(ByVal x As Long, ByVal y As Long) As Long
    ' Clamp x and y values to actual ones
    ' Give one tile worth of leeway before asserting, for units/transporters coming in from off-map.
    If Not (x >= -1) Then ASSERT_FAILURE "x >= -1", "lib/ivis_opengl/pielighting.cpp:36", "operator()", "mapTile: x value is too small (%d,%d) in %dx%d", x, y, mapWidth, mapHeight
    If Not (y >= -1) Then ASSERT_FAILURE "y >= -1", "lib/ivis_opengl/pielighting.cpp:37", "operator()", "mapTile: y value is too small (%d,%d) in %dx%d", x, y, mapWidth, mapHeight
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    If Not (x < mapWidth + 1) Then ASSERT_FAILURE "x < mapWidth + 1", "lib/ivis_opengl/pielighting.cpp:40", "operator()", "mapTile: x value is too big (%d,%d) in %dx%d", x, y, mapWidth, mapHeight
    If Not (y < mapHeight + 1) Then ASSERT_FAILURE "y < mapHeight + 1", "lib/ivis_opengl/pielighting.cpp:41", "operator()", "mapTile: y value is too big (%d,%d) in %dx%d", x, y, mapWidth, mapHeight
    If x > mapWidth - 1 Then x = mapWidth - 1
    If y > mapHeight - 1 Then y = mapHeight - 1

    LightMap_index = x + (y * mapWidth)
End Function

' const PIELIGHT& LightMap::operator()(int32_t x, int32_t y) const
Private Function LightMap_index_const(ByVal x As Long, ByVal y As Long) As Long
    ' Clamp x and y values to actual ones
    ' Give one tile worth of leeway before asserting, for units/transporters coming in from off-map.
    If Not (x >= -1) Then ASSERT_FAILURE "x >= -1", "lib/ivis_opengl/pielighting.cpp:52", "operator()", "mapTile: x value is too small (%d,%d) in %dx%d", x, y, mapWidth, mapHeight
    If Not (y >= -1) Then ASSERT_FAILURE "y >= -1", "lib/ivis_opengl/pielighting.cpp:53", "operator()", "mapTile: y value is too small (%d,%d) in %dx%d", x, y, mapWidth, mapHeight
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    If Not (x < mapWidth + 1) Then ASSERT_FAILURE "x < mapWidth + 1", "lib/ivis_opengl/pielighting.cpp:56", "operator()", "mapTile: x value is too big (%d,%d) in %dx%d", x, y, mapWidth, mapHeight
    If Not (y < mapHeight + 1) Then ASSERT_FAILURE "y < mapHeight + 1", "lib/ivis_opengl/pielighting.cpp:57", "operator()", "mapTile: y value is too big (%d,%d) in %dx%d", x, y, mapWidth, mapHeight
    If x > mapWidth - 1 Then x = mapWidth - 1
    If y > mapHeight - 1 Then y = mapHeight - 1

    LightMap_index_const = x + (y * mapWidth)
End Function

' (VB) getCurrentLightmapData()(x, y), read
Public Function LightMap_get(ByVal x As Long, ByVal y As Long) As Long
    LightMap_get = lightMapData(LightMap_index(x, y))
End Function

' (VB) the const operator(), read (through a const LightMap&)
Public Function LightMap_get_const(ByVal x As Long, ByVal y As Long) As Long
    LightMap_get_const = lightMapData(LightMap_index_const(x, y))
End Function

' (VB) getCurrentLightmapData()(x, y) = colour
Public Sub LightMap_set(ByVal x As Long, ByVal y As Long, ByVal colour As Long)
    lightMapData(LightMap_index(x, y)) = colour
End Sub

' getCurrentLightmapData().reset(width, height)
Public Sub LightMap_reset(ByVal width As Long, ByVal height As Long)
    mapWidth = width
    mapHeight = height
    If width * height > 0 Then
        ReDim lightMapData(0 To width * height - 1)
    Else
        ReDim lightMapData(0 To 0)          ' (VB) std::make_unique<PIELIGHT[]>(0)
    End If
End Sub

' (VB) the light map's PIELIGHTs (mapWidth x mapHeight Longs, row after row), for bulk reads and writes
Public Function LightMap_dataAddress() As Long
    LightMap_dataAddress = VarPtr(lightMapData(0))
End Function

'--------------------------------------------------------------------------
' LightingData (getCurrentLightingData().lights)

Public Sub LightingData_lights_clear()
    lights_size = 0
End Sub

Public Sub LightingData_lights_push_back(light_ As LIGHT)
    If lights_size = 0 Then
        ReDim lights(0 To 63)
    ElseIf lights_size > UBound(lights) Then
        ReDim Preserve lights(0 To 2 * lights_size - 1)
    End If
    lights(lights_size) = light_
    lights_size = lights_size + 1
End Sub

' (VB) for the exe: the LIGHT as its fields
Public Sub LightingData_lights_push_back_fields(ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal range As Long, ByVal colour As Long)
    Dim l As LIGHT
    l.position.x = x
    l.position.y = y
    l.position.z = z
    l.range = range
    l.colour = colour
    LightingData_lights_push_back l
End Sub

Public Function LightingData_lights_size() As Long
    LightingData_lights_size = lights_size
End Function

Public Function LightingData_lights_at(ByVal i As Long) As LIGHT
    LightingData_lights_at = lights(i)
End Function

' (VB) for the exe: lights[i]'s fields
Public Sub LightingData_lights_at_fields(ByVal i As Long, x As Long, y As Long, z As Long, range As Long, colour As Long)
    x = lights(i).position.x
    y = lights(i).position.y
    z = lights(i).position.z
    range = lights(i).range
    colour = lights(i).colour
End Sub

'--------------------------------------------------------------------------

Private Function getLightBoundingBox(light_ As LIGHT) As BoundingBox
    Dim center As Vector3f, range As Single, horizontal As Vector3f, vertical As Vector3f, forward As Vector3f
    Dim horizRange As Vector3f, verticalRange As Vector3f, forwardRange As Vector3f
    Dim centerMinusHorizRange As Vector3f, centerPlusHorizRange As Vector3f
    Dim verticalRangeMinusForwardRange As Vector3f, verticalRangePlusForwardRange As Vector3f
    Dim r As BoundingBox
    center = glm_vec3(light_.position.x, light_.position.y, light_.position.z)
    center.z = center.z * -1!
    range = CSng(U32(light_.range))
    horizontal = glm_vec3(1!, 0!, 0!)
    vertical = glm_vec3(0!, 1!, 0!)
    forward = glm_vec3(0!, 0!, 1!)

    horizRange = glm_vec3_mul(horizontal, range)
    verticalRange = glm_vec3_mul(vertical, range)
    forwardRange = glm_vec3_mul(forward, range)

    centerMinusHorizRange = glm_vec3_sub(center, horizRange)
    centerPlusHorizRange = glm_vec3_add(center, horizRange)
    verticalRangeMinusForwardRange = glm_vec3_sub(verticalRange, forwardRange)
    verticalRangePlusForwardRange = glm_vec3_add(verticalRange, forwardRange)

    r.v(0) = glm_vec3_sub(centerMinusHorizRange, verticalRangeMinusForwardRange)
    r.v(1) = glm_vec3_sub(centerMinusHorizRange, verticalRangePlusForwardRange)
    r.v(2) = glm_vec3_add(centerMinusHorizRange, verticalRangeMinusForwardRange)
    r.v(3) = glm_vec3_add(centerMinusHorizRange, verticalRangePlusForwardRange)
    r.v(4) = glm_vec3_sub(centerPlusHorizRange, verticalRangeMinusForwardRange)
    r.v(5) = glm_vec3_sub(centerPlusHorizRange, verticalRangePlusForwardRange)
    r.v(6) = glm_vec3_add(centerPlusHorizRange, verticalRangeMinusForwardRange)
    r.v(7) = glm_vec3_add(centerPlusHorizRange, verticalRangePlusForwardRange)
    getLightBoundingBox = r
End Function


Private Function pielight_maptile_coord(ByVal worldCoord As Long) As Long
    pielight_maptile_coord = SAR(worldCoord, TILE_SHIFT)
End Function

Private Function pointLightDistanceCalc(a As CalculatedPointLight, b As LIGHT) As Single
    Dim pointLightVector As Vector3f, length As Single
    pointLightVector = glm_vec3_sub(a.position, glm_vec3(b.position.x, b.position.y, b.position.z))
    length = glm_length3(pointLightVector)
    pointLightDistanceCalc = length
End Function

' renderingNew::LightingManager::ComputeFrameData: the class's body (renderingNew_LightingManager.cls calls it
' with its cached culledLights)
' CulledLightInfo: culledLights_light(i) and culledLights_box(i), i < culledLights_size
Public Sub renderingNew_LightingManager_ComputeFrameData(culledLights_light() As CalculatedPointLight, culledLights_box() As BoundingBox, culledLights_size As Long, worldViewProjectionMatrix As mat4)
    Dim result As PointLightBuckets, yAxisInverted As Boolean, viewFrustum As IntersectionOfHalfSpace
    Const maxRangedLightsPerTile As Long = 16
    Const minLightRange As Long = 5
    Const distanceCalcCombineThreshold As Single = 32!
    Dim lightsCombined As Long, lightsSkipped As Long, tinyLightsSkipped As Long
    ' std::unordered_map<std::pair<int32_t, int32_t>, std::vector<size_t>, TileCoordsHasher> tileRangeLights;
    ' (VB) map tile coordinates to vector of culledLight indexes: a list searched in order (at most max_lights entries)
    Dim tileX() As Long, tileY() As Long, tileLights() As Long, tileLights_size() As Long, tiles_size As Long
    Dim li As Long, clipSpaceBoundingBox As BoundingBox, lightTileX As Long, lightTileY As Long, it As Long
    Dim combinedLight As Boolean, k As Long, o As Long, newLightRange As Single, distanceCalc As Single
    Dim calcLight As CalculatedPointLight, weight As Single, lightIndex As Long
    Dim overallId As Long, bucketId As Long, bucketDimension As Long
    Const minBucketDimension As Long = 4
    Dim lightList(0 To gfx_api_max_indexed_lights * 4 - 1) As Long
    Dim reduceNumberOfBucketsNeeded As Boolean, i As Long, j As Long, z As Long
    Dim bucketFrustumX0 As Single, bucketFrustumX1 As Single, bucketFrustumY0 As Single, bucketFrustumY1 As Single
    Dim frustum As IntersectionOfHalfSpace, bucketSize As Long

    result.bucketDimensionUsed = gfx_api_bucket_dimension
    yAxisInverted = gfx_api_context_isYAxisInverted()

    ' Pick the first lights inside the view frustum
    viewFrustum.checks(0) = HalfSpaceCheck_new(HalfSpace_x_ge, -1!)
    viewFrustum.checks(1) = HalfSpaceCheck_new(HalfSpace_x_le, 1!)
    viewFrustum.checks(2) = HalfSpaceCheck_new(HalfSpace_y_ge, -1!, yAxisInverted)
    viewFrustum.checks(3) = HalfSpaceCheck_new(HalfSpace_y_le, 1!, yAxisInverted)
    viewFrustum.checks(4) = HalfSpaceCheck_new(HalfSpace_z_ge, 0!)
    viewFrustum.checks(5) = HalfSpaceCheck_new(HalfSpace_z_le, 1!)

    ReDim tileX(0 To gfx_api_max_lights - 1)
    ReDim tileY(0 To gfx_api_max_lights - 1)
    ReDim tileLights(0 To gfx_api_max_lights - 1, 0 To maxRangedLightsPerTile - 1)
    ReDim tileLights_size(0 To gfx_api_max_lights - 1)
    lightsCombined = 0
    lightsSkipped = 0
    tinyLightsSkipped = 0

    culledLights_size = 0               ' culledLights.clear()
    For li = 0 To lights_size - 1
        If culledLights_size >= gfx_api_max_lights Then
            Exit For
        End If
        clipSpaceBoundingBox = transformBoundingBox(worldViewProjectionMatrix, getLightBoundingBox(lights(li)))
        If Not isBBoxInClipSpace(viewFrustum, clipSpaceBoundingBox) Then
            GoTo continue_
        End If

        If U32(lights(li).range) >= minLightRange Then
            lightTileX = pielight_maptile_coord(lights(li).position.x)
            lightTileY = pielight_maptile_coord(lights(li).position.y)
            ' auto it = tileRangeLights.find(lightTileCoords);
            it = -1
            For k = 0 To tiles_size - 1
                If tileX(k) = lightTileX And tileY(k) = lightTileY Then
                    it = k
                    Exit For
                End If
            Next
            If it <> -1 Then
                ' merge point lights (if possible)
                combinedLight = False
                For k = 0 To tileLights_size(it) - 1
                    o = tileLights(it, k)
                    newLightRange = CSng(U32(lights(li).range))
                    distanceCalc = pointLightDistanceCalc(culledLights_light(o), lights(li))
                    If (distanceCalc < distanceCalcCombineThreshold) And (distanceCalc < (culledLights_light(o).range + newLightRange)) Then
                        ' Found two lights close to each other - combine them
                        If newLightRange > culledLights_light(o).range Then
                            ' If the new light has a greater range, use that as the "base"
                            calcLight.position = glm_vec3(lights(li).position.x, lights(li).position.y, lights(li).position.z)
                            calcLight.colour = glm_vec3(PIELIGHT_r(lights(li).colour) / 255!, PIELIGHT_g(lights(li).colour) / 255!, PIELIGHT_b(lights(li).colour) / 255!)
                            calcLight.range = CSng(U32(lights(li).range))

                            weight = culledLights_light(o).range / calcLight.range
                            calcLight.colour.x = calcLight.colour.x + (culledLights_light(o).colour.x) * weight
                            calcLight.colour.y = calcLight.colour.y + (culledLights_light(o).colour.y) * weight
                            calcLight.colour.z = calcLight.colour.z + (culledLights_light(o).colour.z) * weight

                            culledLights_light(o) = calcLight
                            culledLights_box(o) = clipSpaceBoundingBox
                        Else
                            weight = CSng(U32(lights(li).range)) / culledLights_light(o).range
                            culledLights_light(o).colour.x = culledLights_light(o).colour.x + (PIELIGHT_r(lights(li).colour) / 255!) * weight
                            culledLights_light(o).colour.y = culledLights_light(o).colour.y + (PIELIGHT_g(lights(li).colour) / 255!) * weight
                            culledLights_light(o).colour.z = culledLights_light(o).colour.z + (PIELIGHT_b(lights(li).colour) / 255!) * weight
                        End If
                        combinedLight = True
                        Exit For
                    End If
                Next
                If combinedLight Then
                    lightsCombined = lightsCombined + 1
                    GoTo continue_
                End If
                If tileLights_size(it) >= maxRangedLightsPerTile Then
                    lightsSkipped = lightsSkipped + 1
                    GoTo continue_
                End If

                tileLights(it, tileLights_size(it)) = culledLights_size
                tileLights_size(it) = tileLights_size(it) + 1
            Else
                tileX(tiles_size) = lightTileX
                tileY(tiles_size) = lightTileY
                tileLights(tiles_size, 0) = culledLights_size
                tileLights_size(tiles_size) = 1
                tiles_size = tiles_size + 1
            End If
        Else
            tinyLightsSkipped = tinyLightsSkipped + 1
            GoTo continue_
        End If

        calcLight.position = glm_vec3(lights(li).position.x, lights(li).position.y, lights(li).position.z)
        calcLight.colour = glm_vec3(PIELIGHT_r(lights(li).colour) / 255!, PIELIGHT_g(lights(li).colour) / 255!, PIELIGHT_b(lights(li).colour) / 255!)
        calcLight.range = CSng(U32(lights(li).range))

        culledLights_light(culledLights_size) = calcLight
        culledLights_box(culledLights_size) = clipSpaceBoundingBox
        culledLights_size = culledLights_size + 1
continue_:
    Next

    If lightsSkipped > 0 Or lightsCombined > 0 Or tinyLightsSkipped > 0 Then
        ' debug(LOG_INFO, "Point lights - merged: %zu, skipped (tile limit): %zu, skipped (tiny): %zu", lightsCombined, lightsSkipped, tinyLightsSkipped);
    End If

    For lightIndex = 0 To culledLights_size - 1
        result.positions(lightIndex).x = culledLights_light(lightIndex).position.x
        result.positions(lightIndex).y = culledLights_light(lightIndex).position.y
        result.positions(lightIndex).z = culledLights_light(lightIndex).position.z
        result.colorAndEnergy(lightIndex).x = culledLights_light(lightIndex).colour.x
        result.colorAndEnergy(lightIndex).y = culledLights_light(lightIndex).colour.y
        result.colorAndEnergy(lightIndex).z = culledLights_light(lightIndex).colour.z
        result.colorAndEnergy(lightIndex).w = culledLights_light(lightIndex).range
    Next

    ' Iterate over all buckets
    overallId = 0
    bucketId = 0

    bucketDimension = gfx_api_bucket_dimension ' start off at the maximum number of buckets

    ' GLSL std layout 140 force us to store array of int with the same stride as
    ' an array of ivec4, wasting 3/4 of the storage.
    ' To circumvent this, we pack 4 consecutives index in a ivec4 here, and unpack the value in the shader.
    reduceNumberOfBucketsNeeded = False
    Do
        overallId = 0
        bucketId = 0
        reduceNumberOfBucketsNeeded = False
        For i = 0 To bucketDimension - 1
            bucketFrustumX0 = -1! + 2 * CSng(i) / bucketDimension
            bucketFrustumX1 = -1! + 2 * CSng(i + 1) / bucketDimension

            For j = 0 To bucketDimension - 1
                bucketFrustumY0 = -1! + 2 * CSng(j) / bucketDimension
                bucketFrustumY1 = -1! + 2 * CSng(j + 1) / bucketDimension

                frustum.checks(0) = HalfSpaceCheck_new(HalfSpace_x_ge, bucketFrustumX0)
                frustum.checks(1) = HalfSpaceCheck_new(HalfSpace_x_le, bucketFrustumX1)
                frustum.checks(2) = HalfSpaceCheck_new(HalfSpace_y_ge, bucketFrustumY0, yAxisInverted)
                frustum.checks(3) = HalfSpaceCheck_new(HalfSpace_y_le, bucketFrustumY1, yAxisInverted)
                frustum.checks(4) = HalfSpaceCheck_new(HalfSpace_z_ge, 0!)
                frustum.checks(5) = HalfSpaceCheck_new(HalfSpace_z_le, 1!)

                bucketSize = 0
                For lightIndex = 0 To culledLights_size - 1
                    If overallId + bucketSize >= gfx_api_max_indexed_lights * 4 Then
                        ' number of indexed lights will exceed max permitted - too many buckets
                        reduceNumberOfBucketsNeeded = True
                        Exit For
                    End If

                    If isBBoxInClipSpace(frustum, culledLights_box(lightIndex)) Then
                        lightList(overallId + bucketSize) = lightIndex

                        bucketSize = bucketSize + 1
                    End If
                Next

                result.bucketOffsetAndSize(bucketId) = glm_ivec4(overallId, bucketSize, 0, 0)
                overallId = overallId + bucketSize
                bucketId = bucketId + 1

                If reduceNumberOfBucketsNeeded Then
                    Exit For
                End If
            Next

            If reduceNumberOfBucketsNeeded Then
                If bucketDimension > minBucketDimension Then
                    bucketDimension = bucketDimension - 1
                Else
                    ' reached minimum number of buckets, but still hit max indexed point lights
                    reduceNumberOfBucketsNeeded = False

                    ' zero out remaining buckets
                    ' (note: will mean no point lights for those parts of the screen, but at least we won't have garbage data used)
                    For z = bucketId To gfx_api_bucket_dimension * gfx_api_bucket_dimension - 1
                        result.bucketOffsetAndSize(z) = glm_ivec4(overallId, 0, 0, 0)
                    Next
                End If
                Exit For
            End If
        Next
    Loop While reduceNumberOfBucketsNeeded

    ' pack the index
    For i = 0 To gfx_api_max_indexed_lights * 4 - 1
        Select Case i Mod 4
        Case 0: result.light_index(i \ 4).x = lightList(i)
        Case 1: result.light_index(i \ 4).y = lightList(i)
        Case 2: result.light_index(i \ 4).z = lightList(i)
        Case 3: result.light_index(i \ 4).w = lightList(i)
        End Select
    Next

    result.bucketDimensionUsed = bucketDimension

    currentPointLightBuckets = result
End Sub

'--------------------------------------------------------------------------
' ILightingManager's non-virtual members (the buckets are the current manager's: see the header)

' void SetFrameStart() { currentPointLightBuckets = {}; }
Public Sub ILightingManager_SetFrameStart(ByVal manager As ILightingManager)
    Dim empty_ As PointLightBuckets
    empty_.bucketDimensionUsed = gfx_api_bucket_dimension
    currentPointLightBuckets = empty_
End Sub

Public Function ILightingManager_getPointLightBuckets(ByVal manager As ILightingManager) As PointLightBuckets
    ILightingManager_getPointLightBuckets = currentPointLightBuckets
End Function

' (VB) getPointLightBuckets() for the exe: the four arrays are copied to the addresses (VarPtr of the uniform
' struct's PointLightsPosition(0), PointLightsColorAndEnergy(0), bucketOffsetAndSize(0), indexed_lights(0));
' returns bucketDimensionUsed
Public Function ILightingManager_getPointLightBuckets_p(ByVal manager As ILightingManager, ByVal pPositions As Long, ByVal pColorAndEnergy As Long, ByVal pBucketOffsetAndSize As Long, ByVal pLightIndex As Long) As Long
    CopyMemory ByVal pPositions, currentPointLightBuckets.positions(0), gfx_api_max_lights * 16
    CopyMemory ByVal pColorAndEnergy, currentPointLightBuckets.colorAndEnergy(0), gfx_api_max_lights * 16
    CopyMemory ByVal pBucketOffsetAndSize, currentPointLightBuckets.bucketOffsetAndSize(0), gfx_api_bucket_dimension * gfx_api_bucket_dimension * 16
    CopyMemory ByVal pLightIndex, currentPointLightBuckets.light_index(0), gfx_api_max_indexed_lights * 16
    ILightingManager_getPointLightBuckets_p = currentPointLightBuckets.bucketDimensionUsed
End Function

Public Sub setLightingManager(ByVal manager As ILightingManager)
    Set lightingManager = manager
End Sub

Public Function getCurrentLightingManager() As ILightingManager
    Set getCurrentLightingManager = lightingManager
End Function

' (VB) std::make_unique<renderingNew::LightingManager>()
Public Function renderingNew_LightingManager_new() As ILightingManager
    Set renderingNew_LightingManager_new = New renderingNew_LightingManager
End Function
