Attribute VB_Name = "edit3d_c"
' src/edit3d.cpp, edit3d.h: the structure placement state (and the map editing functions). The
' enum BuildState is BuildState_ (VB names ignore case: the variable buildState), struct
' HIGHLIGHT is HIGHLIGHT_ (WIDGET's highlight). A BUILDCALLBACK
' is a Long, 0 = nullptr: 4.7.0 has no callback (every init3DBuilding passes nullptr), so none is
' ever called (process3DBuilding skips the call).
Option Explicit

'/*
'Edit3D.c - to ultimately contain the map editing functions -
'they are presently scattered in various files .
'Alex McLean, Pumpkin Studios, EIDOS Interactive, 1997
'*/

Public Const TILE_RAISE As Long = 1
Public Const TILE_LOWER As Long = -1

Public Enum BuildState_
    BUILD3D_NONE
    BUILD3D_POS
    BUILD3D_FINISHED
    BUILD3D_VALID
End Enum

Public Type HIGHLIGHT_
    xTL As Long                         ' UWORD: Top left of box to highlight
    yTL As Long
    xBR As Long                         ' UWORD: Bottom right of box to highlight
    yBR As Long
End Type

Public Type BUILDDETAILS
    CallBack As Long                    ' BUILDCALLBACK (0 = nullptr)
    UserData As Long                    ' void *: this holds the OBJECT_POSITION pointer for a Deliv Point
    x As Long                           ' UDWORD
    y As Long
    width As Long
    height As Long
    psStats As Long                     ' BASE_STATS * (0 = nullptr)
    directionShift As Long              ' uint16_t
End Type

'/*
'Definition of a tile to highlight - presently more than is required
'but means that we can highlight any individual tile in future. An
'x coordinate that is greater than mapWidth implies that the highlight
'is invalid (not currently being used)
'*/

Public buildState As BuildState_         ' = BUILD3D_NONE
Public buildSite As HIGHLIGHT_
Public sBuildDetails As BUILDDETAILS
Public brushSize As Long                ' = 1 (edit3d_c_init)
Public quickQueueMode As Boolean        ' = false

' (VB) the C++ initialisers of this module's globals (Sub Main calls it)
Public Sub edit3d_c_init()
    buildState = BUILD3D_NONE
    brushSize = 1
    quickQueueMode = False
End Sub

' Initialisation function for statis & globals in this module.
'
Public Sub Edit3DInitVars()
    buildState = BUILD3D_NONE
    brushSize = 1
End Sub

' (VB) MIN of two ints
Private Function minL(ByVal a As Long, ByVal b As Long) As Long
    If b < a Then minL = b Else minL = a
End Function

'/* Raises a tile by a #defined height */
Public Sub raiseTile(ByVal tile3dX As Long, ByVal tile3dY As Long)
    Dim i As Long, j As Long

    If tile3dX < 0 Or tile3dX > mapWidth - 1 Or tile3dY < 0 Or tile3dY > mapHeight - 1 Then
        Exit Sub
    End If
    For i = tile3dX To minL(mapWidth - 1, tile3dX + brushSize)
        For j = tile3dY To minL(mapHeight - 1, tile3dY + brushSize)
            adjustTileHeight mapTile(i, j), TILE_RAISE
            markTileDirty i, j
        Next
    Next
End Sub

'/* Lowers a tile by a #defined height */
Public Sub lowerTile(ByVal tile3dX As Long, ByVal tile3dY As Long)
    Dim i As Long, j As Long

    If tile3dX < 0 Or tile3dX > mapWidth - 1 Or tile3dY < 0 Or tile3dY > mapHeight - 1 Then
        Exit Sub
    End If
    For i = tile3dX To minL(mapWidth - 1, tile3dX + brushSize)
        For j = tile3dY To minL(mapHeight - 1, tile3dY + brushSize)
            adjustTileHeight mapTile(i, j), TILE_LOWER
            markTileDirty i, j
        Next
    Next
End Sub

'/* Ensures any adjustment to tile elevation is within allowed ranges */
' psTile: a MAPTILE * (its index in psMapTiles)
Public Sub adjustTileHeight(ByVal psTile As Long, ByVal adjust As Long)
    Dim newHeight As Long
    newHeight = psMapTiles(psTile).height + adjust

    If newHeight >= TILE_MIN_HEIGHT And newHeight <= TILE_MAX_HEIGHT Then
        psMapTiles(psTile).height = newHeight
    End If
End Sub

' psStats: a BASE_STATS *; CallBack: a BUILDCALLBACK (0 = nullptr); UserData: a void *
Public Sub init3DBuilding(ByVal psStats As Long, ByVal CallBack As Long, ByVal UserData As Long)
    If Not (psStats <> 0) Then ASSERT_FAILURE "psStats", "src/edit3d.cpp:107", "init3DBuilding", "Bad parameter": Exit Sub

    buildState = BUILD3D_POS

    sBuildDetails.CallBack = CallBack
    sBuildDetails.UserData = UserData

    If BASE_STATS_hasType(psStats, STAT_STRUCTURE) Then
        sBuildDetails.width = asStructureStats(castStructureStats(psStats)).baseWidth
        sBuildDetails.height = asStructureStats(castStructureStats(psStats)).baseBreadth
        sBuildDetails.psStats = psStats
    ElseIf BASE_STATS_hasType(psStats, STAT_FEATURE) Then
        sBuildDetails.width = asFeatureStats(psStats And &HFFFF&).baseWidth
        sBuildDetails.height = asFeatureStats(psStats And &HFFFF&).baseBreadth
        sBuildDetails.psStats = psStats
    Else '/*if (psStats->hasType(STAT_TEMPLATE))*/
        sBuildDetails.width = 1
        sBuildDetails.height = 1
        sBuildDetails.psStats = psStats
    End If

    sBuildDetails.x = map_coord(mousePos.x - world_coord(sBuildDetails.width) \ 2)
    sBuildDetails.y = map_coord(mousePos.y - world_coord(sBuildDetails.height) \ 2)
End Sub

Public Sub kill3DBuilding()
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    'cancel the drag boxes
    dragBox3D.status = DRAG_INACTIVE
    wallDrag.status = DRAG_INACTIVE
    buildState = BUILD3D_NONE
End Sub


' Call once per frame to handle structure positioning and callbacks.
'
Public Function process3DBuilding() As Boolean
    Dim border As Long, bv As Vector2i, size As Vector2i, worldSize As Vector2i, isValid As Boolean
    Dim lb As LineBuild, i As Long, p As Vector2i

    'if not trying to build ignore
    If buildState = BUILD3D_NONE Then
        If quickQueueMode And Not ctrlShiftDown() Then
            quickQueueMode = False
            intDemolishCancel
        End If
        process3DBuilding = True
        Exit Function
    End If

    '/* Need to update the building locations if we're building */
    border = 5 * TILE_UNITS \ 2
    bv.x = clip_(mousePos.x, border, mapWidth * TILE_UNITS - border)
    bv.y = clip_(mousePos.y, border, mapHeight * TILE_UNITS - border)
    size = getStatsSize(sBuildDetails.psStats, getBuildingDirection())
    worldSize.x = world_coord(size.x)
    worldSize.y = world_coord(size.y)
    ' bv = round_to_nearest_tile(bv - worldSize/2) + worldSize/2;
    bv.x = round_to_nearest_tile(bv.x - worldSize.x \ 2) + worldSize.x \ 2
    bv.y = round_to_nearest_tile(bv.y - worldSize.y \ 2) + worldSize.y \ 2

    If buildState <> BUILD3D_FINISHED Then
        isValid = True
        If wallDrag.status <> DRAG_INACTIVE And BASE_STATS_hasType(sBuildDetails.psStats, STAT_STRUCTURE) And canLineBuild() Then
            wallDrag.pos2 = mousePos ' Why must this be done here? If not doing it here, dragging works almost normally, except it suddenly stops working if the drag becomes invalid.

            lb = calcLineBuild(castStructureStats(sBuildDetails.psStats), getBuildingDirection(), wallDrag.pos, wallDrag.pos2)
            i = 0
            Do While i < lb.count And isValid
                p = LineBuild_at(lb, i)
                isValid = isValid And validLocation(sBuildDetails.psStats, p, getBuildingDirection(), selectedPlayer, True)
                i = i + 1
            Loop
        Else
            isValid = validLocation(sBuildDetails.psStats, bv, getBuildingDirection(), selectedPlayer, True)
        End If

        If isValid Then buildState = BUILD3D_VALID Else buildState = BUILD3D_POS
    End If

    buildSite.xTL = map_coord(bv.x - worldSize.x \ 2) And &HFFFF&
    sBuildDetails.x = buildSite.xTL
    buildSite.yTL = map_coord(bv.y - worldSize.y \ 2) And &HFFFF&
    sBuildDetails.y = buildSite.yTL
    If (getBuildingDirection() And &H4000&) = 0 Then
        buildSite.xBR = (buildSite.xTL + sBuildDetails.width - 1) And &HFFFF&
        buildSite.yBR = (buildSite.yTL + sBuildDetails.height - 1) And &HFFFF&
    Else
        ' Rotated 90 degrees, swap width and height
        buildSite.xBR = (buildSite.xTL + sBuildDetails.height - 1) And &HFFFF&
        buildSite.yBR = (buildSite.yTL + sBuildDetails.width - 1) And &HFFFF&
    End If

    If (buildState = BUILD3D_FINISHED) And (sBuildDetails.CallBack <> 0) Then
        ' sBuildDetails.CallBack(sBuildDetails.x, sBuildDetails.y, sBuildDetails.UserData): (VB) no BUILDCALLBACK exists (the module's header)
        buildState = BUILD3D_NONE
        process3DBuilding = True
        Exit Function
    End If
    If quickQueueMode And Not ctrlShiftDown() Then
        buildState = BUILD3D_NONE
        quickQueueMode = False
    End If

    process3DBuilding = False
End Function

' amount: uint16_t
Public Sub incrementBuildingDirection(ByVal amount As Long)
    sBuildDetails.directionShift = (sBuildDetails.directionShift + amount) And &HFFFF&
End Sub

Public Function getBuildingDirection() As Long
    getBuildingDirection = snapDirection(playerPos.r.y + sBuildDetails.directionShift)
End Function

'/* See if a structure location has been found */
Public Function found3DBuilding(pos As Vector2i) As Boolean
    If buildState <> BUILD3D_FINISHED Then
        found3DBuilding = False
        Exit Function
    End If

    ' pos = world_coord({sBuildDetails.x, sBuildDetails.y}) + world_coord((getBuildingDirection() & 0x4000) == 0? {width, height} : {height, width})/2;
    If (getBuildingDirection() And &H4000&) = 0 Then
        pos.x = world_coord(sBuildDetails.x) + world_coord(sBuildDetails.width) \ 2
        pos.y = world_coord(sBuildDetails.y) + world_coord(sBuildDetails.height) \ 2
    Else
        pos.x = world_coord(sBuildDetails.x) + world_coord(sBuildDetails.height) \ 2
        pos.y = world_coord(sBuildDetails.y) + world_coord(sBuildDetails.width) \ 2
    End If

    If ctrlShiftDown() Then
        quickQueueMode = True
        init3DBuilding sBuildDetails.psStats, sBuildDetails.CallBack, sBuildDetails.UserData
    Else
        buildState = BUILD3D_NONE
    End If

    found3DBuilding = True
End Function

'/* See if a second position for a build has been found */
Public Function found3DBuildLocTwo(pos As Vector2i, pos2 As Vector2i) As Boolean
    Dim stats As Long, lb As LineBuild
    If wallDrag.status <> DRAG_RELEASED Then
        found3DBuildLocTwo = False
        Exit Function
    End If
    If Not canLineBuild() Then
        found3DBuildLocTwo = False
        Exit Function
    End If

    'whilst we're still looking for a valid location - return false
    If buildState = BUILD3D_POS Then
        found3DBuildLocTwo = False
        Exit Function
    End If

    wallDrag.status = DRAG_INACTIVE
    stats = castStructureStats(sBuildDetails.psStats)
    lb = calcLineBuild(stats, getBuildingDirection(), wallDrag.pos, wallDrag.pos2)
    pos = lb.begin_
    pos2 = LineBuild_back(lb)

    If ctrlShiftDown() Then
        quickQueueMode = True
        init3DBuilding sBuildDetails.psStats, sBuildDetails.CallBack, sBuildDetails.UserData
    End If

    found3DBuildLocTwo = True
End Function

'/*returns true if the build state is not equal to BUILD3D_NONE*/
Public Function tryingToGetLocation() As Boolean
    If vbDisplayOff Then tryingToGetLocation = False: Exit Function  ' (VB) the vbtest runs: statsref's stand-in
    If buildState = BUILD3D_NONE Then
        tryingToGetLocation = False
    Else
        tryingToGetLocation = True
    End If
End Function

' edit3d.h: static inline bool canLineBuild()
' May only call if sBuildDetails.psStats points to a STRUCTURE_STATS.
Public Function canLineBuild() As Boolean
    If Not (selectedPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/edit3d.h:77", "canLineBuild", "selectedPlayer definitely can't line-build: %u", selectedPlayer: canLineBuild = False: Exit Function
    canLineBuild = asStructureStats(castStructureStats(sBuildDetails.psStats)).upgrade(selectedPlayer).limit > 1
End Function
