Attribute VB_Name = "visibility_c"
' src/visibility.cpp, visibility.h: the tiles an object sees
' (visTilesUpdate and its wave cast, visRemoveVisibility, setUnderTilesVis,
' revealAll), the spotters (addSpotter, removeSpotter, removeSpotters),
' hasSharedVision, the objects seen each tick (processVisibility,
' visUpdateLevel), visibility.h's inline functions, and the line of sight
' and fire (visibleObject with its ray cast callback rayLOSCallback,
' RAY_CALLBACK_LOS; visGetBlockingWall, lineOfFire, areaOfFire,
' arcOfFire). A BASE_OBJECT's watchedTiles (std::vector<TILEPOS>) is its
' array watchedTiles plus watchedTiles_size, reached through the watched*
' helpers below. The class SPOTTER is a UDT; apsInvisibleViewers
' (std::vector<SPOTTER *>) holds the spotters themselves, its destructor
' is spotterDestruct.
Option Explicit

' accuracy for the height gradient
Private Const GRAD_MUL As Long = 10000

' rate to change visibility level
Private Const VIS_LEVEL_INC As Long = 255 * 2
Private Const VIS_LEVEL_DEC As Long = 50

' integer amount to change visibility this turn
Private visLevelInc As Long
Private visLevelDec As Long

Private Const MIN_VIS_HEIGHT As Long = 80

' Trivial upper bound to what a fully upgraded WSS can use (its number of angles). Should probably be some factor times the maximum possible radius. Is probably a lot more than needed. Tested to need at least 180.
Private Const MAX_WAVECAST_LIST_SIZE As Long = 1360
' doWaveTerrain's int heights[2][MAX_WAVECAST_LIST_SIZE] and size_t angles[2][MAX_WAVECAST_LIST_SIZE + 1]
Private heights(0 To 1, 0 To MAX_WAVECAST_LIST_SIZE - 1) As Long
Private angles(0 To 1, 0 To MAX_WAVECAST_LIST_SIZE) As Long
' visibility.h
Public Const LINE_OF_FIRE_MINIMUM As Long = 5

' visibleObject's ray cast data (VisibleObjectHelp_t), rayLOSCallback's void *data
Private Type VisibleObjectHelp_t
    rayStart As Boolean     ' Whether this is the first point on the ray
    wallsBlock As Boolean   ' Whether walls block line of sight
    startHeight As Long     ' The height at the view point
    final As Vector2i       ' The final tile of the ray cast
    lastHeight As Long      ' The last height and distance
    lastDist As Long
    currGrad As Long        ' The current obscuring gradient
    numWalls As Long        ' Whether the LOS has hit a wall
    wall As Vector2i        ' The position of a wall if it is on the LOS
End Type
Private losHelp As VisibleObjectHelp_t

' static int *gNumWalls; static Vector2i *gWall: visGetBlockingWall's out globals (gWallSet: != nullptr)
Private gWallSet As Boolean
Private gNumWalls As Long
Private gWall As Vector2i

Private Type SPOTTER
    pos As Vector3i                     ' Position (map tiles)
    player As Long
    sensorRadius As Long
    sensorType As Long                  ' 0 - vision, 1 - radar
    expiryTime As Long                  ' uint32_t: when to self-destruct, zero if never
    numWatchedTiles As Long
    watchedTiles() As TILEPOS
    id As Long                          ' uint32_t
End Type
Private apsInvisibleViewers() As SPOTTER    ' std::vector<SPOTTER *>
Private apsInvisibleViewers_size As Long

Private gridListSpotters() As Long      ' updateSpotters' static GridList
Private gridListVision() As Long        ' processVisibilityVision's static GridList

' ---- (VB) a BASE_OBJECT's watchedTiles

Private Sub watchedClear(ByVal psObj As Long)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).watchedTiles_size = 0
    Case OBJ_STRUCTURE: aStruct(psObj).watchedTiles_size = 0
    Case OBJ_FEATURE: aFeature(psObj).watchedTiles_size = 0
    End Select
End Sub

Private Function watchedSize(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: watchedSize = aDroid(psObj).watchedTiles_size
    Case OBJ_STRUCTURE: watchedSize = aStruct(psObj).watchedTiles_size
    Case OBJ_FEATURE: watchedSize = aFeature(psObj).watchedTiles_size
    End Select
End Function

Private Function watchedAt(ByVal psObj As Long, ByVal i As Long) As TILEPOS
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: watchedAt = aDroid(psObj).watchedTiles(i)
    Case OBJ_STRUCTURE: watchedAt = aStruct(psObj).watchedTiles(i)
    Case OBJ_FEATURE: watchedAt = aFeature(psObj).watchedTiles(i)
    End Select
End Function

Private Sub watchedPush(ByVal psObj As Long, tp As TILEPOS)
    Dim n As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID
        n = aDroid(psObj).watchedTiles_size
        If n = 0 Then
            ReDim aDroid(psObj).watchedTiles(0 To 63)
        ElseIf n > UBound(aDroid(psObj).watchedTiles) Then
            ReDim Preserve aDroid(psObj).watchedTiles(0 To n * 2 - 1)
        End If
        aDroid(psObj).watchedTiles(n) = tp
        aDroid(psObj).watchedTiles_size = n + 1
    Case OBJ_STRUCTURE
        n = aStruct(psObj).watchedTiles_size
        If n = 0 Then
            ReDim aStruct(psObj).watchedTiles(0 To 63)
        ElseIf n > UBound(aStruct(psObj).watchedTiles) Then
            ReDim Preserve aStruct(psObj).watchedTiles(0 To n * 2 - 1)
        End If
        aStruct(psObj).watchedTiles(n) = tp
        aStruct(psObj).watchedTiles_size = n + 1
    Case OBJ_FEATURE
        n = aFeature(psObj).watchedTiles_size
        If n = 0 Then
            ReDim aFeature(psObj).watchedTiles(0 To 63)
        ElseIf n > UBound(aFeature(psObj).watchedTiles) Then
            ReDim Preserve aFeature(psObj).watchedTiles(0 To n * 2 - 1)
        End If
        aFeature(psObj).watchedTiles(n) = tp
        aFeature(psObj).watchedTiles_size = n + 1
    End Select
End Sub

' initialise the visibility stuff
Public Function visInitialise() As Boolean
    visLevelInc = 1
    visLevelDec = 0

    visInitialise = True
End Function

Private Sub updateTileVis(ByVal psTile As Long, ByVal player As Long)
    ' The definition of whether a player can see something on a given tile or not
    Dim seen As Boolean
    If psMapTiles(psTile).watchers(player) > 0 Then
        seen = True
    ElseIf psMapTiles(psTile).sensors(player) > 0 Then
        seen = (psMapTiles(psTile).jammerBits And Not alliancebits(player)) = 0
    End If
    If seen Then
        psMapTiles(psTile).sensorBits = psMapTiles(psTile).sensorBits Or BIT_(player)           ' mark it as being seen
    Else
        psMapTiles(psTile).sensorBits = psMapTiles(psTile).sensorBits And Not BIT_(player)      ' mark as hidden
    End If
End Sub

' update the visibility change levels
Public Sub visUpdateLevel()
    visLevelInc = gameTimeAdjustedAverage(VIS_LEVEL_INC)
    visLevelDec = gameTimeAdjustedAverage(VIS_LEVEL_DEC)
End Sub

Public Function addSpotter(ByVal x As Long, ByVal y As Long, ByVal player As Long, ByVal radius As Long, ByVal radar As Boolean, Optional ByVal expiry As Long = 0) As Long
    Dim psSpot As SPOTTER, size As Long, tbl As Long, i As Long, mapX As Long, mapY As Long, psTile As Long
    Dim tilePos_ As TILEPOS, t As WavecastTile, below As Boolean

    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/visibility.cpp:135", "addSpotter", "invalid player: %d", player: addSpotter = 0: Exit Function
    ' new SPOTTER(x, y, player, radius, (int)radar, expiry)
    psSpot.pos = Vector3i_(x, y, 0)
    psSpot.player = player
    psSpot.sensorRadius = radius
    psSpot.sensorType = -CLng(radar)
    psSpot.expiryTime = expiry
    psSpot.numWatchedTiles = 0
    psSpot.id = generateSynchronisedObjectId()
    tbl = getWavecastTable(radius, size)
    If size > 0 Then ReDim psSpot.watchedTiles(0 To size - 1)
    For i = 0 To size - 1
        t = wavecastTables(tbl).tiles(i)
        mapX = x + t.dx
        mapY = y + t.dy
        If mapX < 0 Or mapX >= mapWidth Or mapY < 0 Or mapY >= mapHeight Then
            GoTo continue_
        End If
        psTile = mapTile(mapX, mapY)
        psMapTiles(psTile).tileExploredBits = psMapTiles(psTile).tileExploredBits Or alliancebits(player)
        If Not radar Then
            below = psMapTiles(psTile).watchers(player) < UWORD_MAX
        Else
            below = psMapTiles(psTile).sensors(player) < UWORD_MAX
        End If
        If below Then
            tilePos_.x = mapX And &HFF&
            tilePos_.y = mapY And &HFF&
            tilePos_.typ = -CLng(radar)
            ' we observe this tile
            If Not radar Then
                psMapTiles(psTile).watchers(player) = psMapTiles(psTile).watchers(player) + 1
            Else
                psMapTiles(psTile).sensors(player) = psMapTiles(psTile).sensors(player) + 1
            End If
            updateTileVis psTile, player
            psSpot.watchedTiles(psSpot.numWatchedTiles) = tilePos_  ' record having seen it
            psSpot.numWatchedTiles = psSpot.numWatchedTiles + 1
        End If
continue_:
    Next
    ' apsInvisibleViewers.push_back(psSpot)
    If apsInvisibleViewers_size = 0 Then
        ReDim apsInvisibleViewers(0 To 7)
    ElseIf apsInvisibleViewers_size > UBound(apsInvisibleViewers) Then
        ReDim Preserve apsInvisibleViewers(0 To 2 * apsInvisibleViewers_size - 1)
    End If
    apsInvisibleViewers(apsInvisibleViewers_size) = psSpot
    apsInvisibleViewers_size = apsInvisibleViewers_size + 1
    addSpotter = psSpot.id
End Function

' (VB) apsInvisibleViewers.erase(apsInvisibleViewers.begin() + i)
Private Sub spotterErase(ByVal i As Long)
    Dim j As Long
    For j = i To apsInvisibleViewers_size - 2
        apsInvisibleViewers(j) = apsInvisibleViewers(j + 1)
    Next
    apsInvisibleViewers_size = apsInvisibleViewers_size - 1
End Sub

Public Function removeSpotter(ByVal id As Long) As Boolean
    Dim i As Long
    For i = 0 To apsInvisibleViewers_size - 1
        If apsInvisibleViewers(i).id = id Then
            spotterDestruct i
            spotterErase i
            removeSpotter = True
            Exit Function
        End If
    Next
    removeSpotter = False
End Function

Public Sub removeSpotters()
    Do While apsInvisibleViewers_size > 0
        spotterDestruct apsInvisibleViewers_size - 1
        apsInvisibleViewers_size = apsInvisibleViewers_size - 1
    Loop
End Sub

Private Sub updateSpotters()
    Dim i As Long, n As Long, gi As Long
    i = 0
    Do While i < apsInvisibleViewers_size
        If apsInvisibleViewers(i).expiryTime <> 0 Then
            If ULT(apsInvisibleViewers(i).expiryTime, gameTime) Then
                spotterDestruct i
                spotterErase i
                GoTo continue_
            End If
        End If
        ' else, ie if not expired, show objects around it
        n = gridStartIterateUnseen(world_coord(apsInvisibleViewers(i).pos.x), world_coord(apsInvisibleViewers(i).pos.y), apsInvisibleViewers(i).sensorRadius, apsInvisibleViewers(i).player, gridListSpotters)
        For gi = 0 To n - 1
            ' Tell system that this side can see this object
            setSeenBy gridListSpotters(gi), apsInvisibleViewers(i).player, UBYTE_MAX
        Next
continue_:
        i = i + 1
    Loop
End Sub

' SPOTTER::~SPOTTER() of apsInvisibleViewers(i)
Private Sub spotterDestruct(ByVal i As Long)
    Dim j As Long, tilePos_ As TILEPOS, psTile As Long, player As Long
    player = apsInvisibleViewers(i).player
    For j = 0 To apsInvisibleViewers(i).numWatchedTiles - 1
        tilePos_ = apsInvisibleViewers(i).watchedTiles(j)
        psTile = mapTile(tilePos_.x, tilePos_.y)
        If tilePos_.typ = 0 Then
            If Not (psMapTiles(psTile).watchers(player) > 0) Then ASSERT_FAILURE "visionType[player] > 0", "src/visibility.cpp:219", "~SPOTTER", "Not watching watched tile (%d, %d)", tilePos_.x, tilePos_.y
            psMapTiles(psTile).watchers(player) = (psMapTiles(psTile).watchers(player) - 1) And &HFFFF&
        Else
            If Not (psMapTiles(psTile).sensors(player) > 0) Then ASSERT_FAILURE "visionType[player] > 0", "src/visibility.cpp:219", "~SPOTTER", "Not watching watched tile (%d, %d)", tilePos_.x, tilePos_.y
            psMapTiles(psTile).sensors(player) = (psMapTiles(psTile).sensors(player) - 1) And &HFFFF&
        End If
        updateTileVis psTile, player
    Next
    Erase apsInvisibleViewers(i).watchedTiles
    apsInvisibleViewers(i).numWatchedTiles = 0
End Sub

' Record all tiles that some object confers visibility to. Only record each tile
' once. Note that there is both a limit to how many objects can watch any given
' tile. Strange but non fatal things will happen if these limits are exceeded.
Private Sub visMarkTile(ByVal psObj As Long, ByVal mapX As Long, ByVal mapY As Long, ByVal psTile As Long)
    Dim rayPlayer As Long, xdiff As Long, ydiff As Long, distSq As Long, inRange As Boolean, tp As TILEPOS
    Dim pos As Vector3i
    rayPlayer = SIMPLE_OBJECT_player(psObj)
    pos = SIMPLE_OBJECT_pos(psObj)
    xdiff = map_coord(pos.x) - mapX
    ydiff = map_coord(pos.y) - mapY
    distSq = xdiff * xdiff + ydiff * ydiff
    inRange = (distSq < 16)

    If inRange Then
        If psMapTiles(psTile).watchers(rayPlayer) >= UWORD_MAX Then Exit Sub
    Else
        If psMapTiles(psTile).sensors(rayPlayer) >= UWORD_MAX Then Exit Sub
    End If
    tp.x = mapX And &HFF&
    tp.y = mapY And &HFF&
    tp.typ = -CLng(inRange)

    ' we observe this tile
    If inRange Then
        psMapTiles(psTile).watchers(rayPlayer) = psMapTiles(psTile).watchers(rayPlayer) + 1
    Else
        psMapTiles(psTile).sensors(rayPlayer) = psMapTiles(psTile).sensors(rayPlayer) + 1
    End If
    If (BASE_OBJECT_flags(psObj) And BIT_(OBJECT_FLAG_JAMMED_TILES)) <> 0 Then  ' we are a jammer object
        psMapTiles(psTile).jammers(rayPlayer) = (psMapTiles(psTile).jammers(rayPlayer) + 1) And &HFFFF&
        psMapTiles(psTile).jammerBits = psMapTiles(psTile).jammerBits Or BIT_(rayPlayer)  ' mark it as being jammed
    End If
    updateTileVis psTile, rayPlayer
    watchedPush psObj, tp  ' record having seen it
End Sub

' The terrain revealing ray callback
Private Sub doWaveTerrain(ByVal psObj As Long)
    Dim pos As Vector3i, sx As Long, sy As Long, sz As Long, radius As Long, rayPlayer As Long, imd As Long
    Dim size As Long, tbl As Long, i As Long
    Dim readListSize As Long, readListPos As Long, writeListPos As Long, readList As Long, lastHeight As Long
    Dim lastAngle As Long, mapX As Long, mapY As Long, psTile As Long, tileHeight As Long
    Dim perspectiveHeight As Long, perspectiveHeightLeeway As Long, seen As Boolean, oldHeight As Long, newHeight As Long
    Dim t As WavecastTile

    If psObj = 0 Then
        Exit Sub
    End If

    pos = SIMPLE_OBJECT_pos(psObj)
    sx = pos.x
    sy = pos.y
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: imd = aDroid(psObj).sDisplay.imd
    Case OBJ_STRUCTURE: imd = aStruct(psObj).sDisplay.imd
    Case OBJ_FEATURE: imd = aFeature(psObj).sDisplay.imd
    End Select
    If imd <> 0 Then sz = pos.z + MAX_(MIN_VIS_HEIGHT, imdMaxY(imd)) Else sz = pos.z + MIN_VIS_HEIGHT
    radius = objSensorRange(psObj)
    rayPlayer = SIMPLE_OBJECT_player(psObj)
    tbl = getWavecastTable(radius, size)
    readListSize = 0: readListPos = 0: writeListPos = 0  ' readListSize, readListPos dummy initialisations.
    readList = 0  ' Reading from this list, writing to the other. Could also initialise to rand()%2.
    lastHeight = 0  ' lastHeight dummy initialisation.
    lastAngle = &H7FFFFFFF  ' std::numeric_limits<size_t>::max()

    ' Start with full vision of all angles. (If someday wanting to make droids that can only look in one direction, change here, after getting the original angle values saved in the wavecast table.)
    heights(1 - readList, writeListPos) = &H80000000    ' Smallest integer.
    angles(1 - readList, writeListPos) = 0              ' Smallest angle.
    writeListPos = writeListPos + 1

    watchedClear psObj
    For i = 0 To size - 1
        t = wavecastTables(tbl).tiles(i)
        mapX = map_coord(sx) + t.dx
        mapY = map_coord(sy) + t.dy
        If mapX < 0 Or mapX >= mapWidth Or mapY < 0 Or mapY >= mapHeight Then
            GoTo continue_
        End If

        psTile = mapTile(mapX, mapY)
        tileHeight = MAX_(psMapTiles(psTile).height, psMapTiles(psTile).waterLevel)  ' If we can see the water surface, then let us see water-covered tiles too.
        perspectiveHeight = (tileHeight - sz) * t.invRadius
        perspectiveHeightLeeway = (tileHeight - sz + MIN_VIS_HEIGHT) * t.invRadius

        If t.angBegin < lastAngle Then
            ' Gone around the circle. (Or just started scan.)
            angles(1 - readList, writeListPos) = lastAngle

            ' Flip the lists.
            readList = 1 - readList
            readListPos = 0
            readListSize = writeListPos
            writeListPos = 0
            lastHeight = 1  ' Impossible value since tiles[i].invRadius > 1 for all i, so triggers writing first entry in list.
        End If
        lastAngle = t.angEnd

        Do While angles(readList, readListPos + 1) <= t.angBegin And readListPos < readListSize
            readListPos = readListPos + 1  ' Skip, not relevant.
        Loop

        seen = False
        Do While angles(readList, readListPos) < t.angEnd And readListPos < readListSize
            oldHeight = heights(readList, readListPos)
            newHeight = MAX_(oldHeight, perspectiveHeight)
            seen = seen Or perspectiveHeightLeeway >= oldHeight ' consider point slightly above ground in case there is something on the tile
            If newHeight <> lastHeight Then
                heights(1 - readList, writeListPos) = newHeight
                angles(1 - readList, writeListPos) = MAX_(angles(readList, readListPos), t.angBegin)
                lastHeight = newHeight
                writeListPos = writeListPos + 1
                If Not (writeListPos <= MAX_WAVECAST_LIST_SIZE) Then ASSERT_FAILURE "writeListPos <= MAX_WAVECAST_LIST_SIZE", "src/visibility.cpp:327", "doWaveTerrain", "Visibility too complicated! Need to increase MAX_WAVECAST_LIST_SIZE.": Exit Sub
            End If
            readListPos = readListPos + 1
        Loop
        readListPos = readListPos - 1

        If seen Then
            ' Can see this tile.
            psMapTiles(psTile).tileExploredBits = psMapTiles(psTile).tileExploredBits Or alliancebits(rayPlayer)  ' Share exploration with allies too
            visMarkTile psObj, mapX, mapY, psTile   ' Mark this tile as seen by our sensor
        End If
continue_:
    Next
End Sub

' Remove tile visibility from object
Public Sub visRemoveVisibility(ByVal psObj As Long)
    Dim i As Long, pos As TILEPOS, psTile As Long, player As Long, jammer As Boolean
    If mapWidth <> 0 And mapHeight <> 0 Then
        player = SIMPLE_OBJECT_player(psObj)
        jammer = (BASE_OBJECT_flags(psObj) And BIT_(OBJECT_FLAG_JAMMED_TILES)) <> 0
        For i = 0 To watchedSize(psObj) - 1
            pos = watchedAt(psObj, i)
            ' FIXME: the mapTile might have been swapped out, see swapMissionPointers()
            psTile = mapTile(pos.x, pos.y)

            If Not (pos.typ < 2) Then ASSERT_FAILURE "pos.type < 2", "src/visibility.cpp:401", "visRemoveVisibility", "Invalid visibility type %d", CLng(pos.typ)
            If pos.typ = 0 Then
                If psMapTiles(psTile).sensors(player) = 0 And game.typ = LEVEL_TYPE_CAMPAIGN Then GoTo continue_  ' hack
                If Not (psMapTiles(psTile).sensors(player) > 0) Then ASSERT_FAILURE "visionType[psObj->player] > 0", "src/visibility.cpp:407", "visRemoveVisibility", "No %s on watched tile (%d, %d)", "vision", CLng(pos.x), CLng(pos.y)
                psMapTiles(psTile).sensors(player) = (psMapTiles(psTile).sensors(player) - 1) And &HFFFF&
            Else
                If psMapTiles(psTile).watchers(player) = 0 And game.typ = LEVEL_TYPE_CAMPAIGN Then GoTo continue_  ' hack
                If Not (psMapTiles(psTile).watchers(player) > 0) Then ASSERT_FAILURE "visionType[psObj->player] > 0", "src/visibility.cpp:407", "visRemoveVisibility", "No %s on watched tile (%d, %d)", "radar", CLng(pos.x), CLng(pos.y)
                psMapTiles(psTile).watchers(player) = (psMapTiles(psTile).watchers(player) - 1) And &HFFFF&
            End If
            If jammer Then  ' we are a jammer object - we cannot check objJammerPower(psObj) > 0 directly here, we may be in the BASE_OBJECT destructor).
                ' No jammers in campaign, no need for special hack
                If Not (psMapTiles(psTile).jammers(player) > 0) Then ASSERT_FAILURE "psTile->jammers[psObj->player] > 0", "src/visibility.cpp:412", "visRemoveVisibility", "Not jamming watched tile (%d, %d)", CLng(pos.x), CLng(pos.y)
                psMapTiles(psTile).jammers(player) = (psMapTiles(psTile).jammers(player) - 1) And &HFFFF&
                If psMapTiles(psTile).jammers(player) = 0 Then
                    psMapTiles(psTile).jammerBits = psMapTiles(psTile).jammerBits And Not BIT_(player)
                End If
            End If
            updateTileVis psTile, player
continue_:
        Next
    End If
    watchedClear psObj
    BASE_OBJECT_flags(psObj) = BASE_OBJECT_flags(psObj) And Not BIT_(OBJECT_FLAG_JAMMED_TILES)
End Sub

Public Sub visRemoveVisibilityOffWorld(ByVal psObj As Long)
    watchedClear psObj
End Sub

' Check which tiles can be seen by an object
Public Sub visTilesUpdate(ByVal psObj As Long)
    If Not (SIMPLE_OBJECT_type(psObj) <> OBJ_FEATURE) Then ASSERT_FAILURE "psObj->type != OBJ_FEATURE", "src/visibility.cpp:434", "visTilesUpdate", "visTilesUpdate: visibility updates are not for features!"

    ' Remove previous map visibility provided by object
    visRemoveVisibility psObj

    If SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE Then
        If aStruct(psObj).status <> SS_BUILT Then Exit Sub
        Select Case asStructureStats(aStruct(psObj).pStructureType).typ
        Case REF_WALL, REF_WALLCORNER, REF_GATE
            ' unbuilt structures and walls do not confer visibility.
            Exit Sub
        End Select
    End If

    ' Do the whole circle in infinite steps. No more pretty moire patterns.
    If objJammerPower(psObj) > 0 Then
        BASE_OBJECT_flags(psObj) = BASE_OBJECT_flags(psObj) Or BIT_(OBJECT_FLAG_JAMMED_TILES)
    Else
        BASE_OBJECT_flags(psObj) = BASE_OBJECT_flags(psObj) And Not BIT_(OBJECT_FLAG_JAMMED_TILES)
    End If
    doWaveTerrain psObj
End Sub

'reveals all the terrain in the map
Public Sub revealAll(ByVal player As Long)
    Dim i As Long, j As Long, psTile As Long

    If player >= MAX_PLAYERS Then
        Exit Sub
    End If

    'reveal all tiles
    For j = 0 To mapHeight - 1
        For i = 0 To mapWidth - 1
            psTile = mapTile(i, j)
            psMapTiles(psTile).tileExploredBits = psMapTiles(psTile).tileExploredBits Or alliancebits(player)
        Next
    Next

    'the objects gets revealed in processVisibility()
End Sub

Public Function hasSharedVision(ByVal viewer As Long, ByVal ally As Long) As Boolean
    If Not ULT(viewer, MAX_PLAYERS) Then hasSharedVision = False: Exit Function
    If Not (ULT(ally, MAX_PLAYERS)) Then ASSERT_FAILURE "ally < MAX_PLAYERS", "src/visibility.cpp:653", "hasSharedVision", "Bad ally %u (viewer: %u)", ally, viewer: Exit Function

    If viewer = ally Then
        hasSharedVision = True
    ElseIf bMultiPlayer Then
        If alliancesSharedVision(game.alliance) Then hasSharedVision = aiCheckAlliances(viewer, ally)
    End If
End Function

' (VB) psObj->seenThisTick[player]
Private Function seenThisTick_(ByVal psObj As Long, ByVal player As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: seenThisTick_ = aDroid(psObj).seenThisTick(player)
    Case OBJ_STRUCTURE: seenThisTick_ = aStruct(psObj).seenThisTick(player)
    Case OBJ_FEATURE: seenThisTick_ = aFeature(psObj).seenThisTick(player)
    End Select
End Function

Private Sub setSeenThisTick(ByVal psObj As Long, ByVal player As Long, ByVal v As Long)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).seenThisTick(player) = v And &HFF&
    Case OBJ_STRUCTURE: aStruct(psObj).seenThisTick(player) = v And &HFF&
    Case OBJ_FEATURE: aFeature(psObj).seenThisTick(player) = v And &HFF&
    End Select
End Sub

Private Sub setSeenBy(ByVal psObj As Long, ByVal viewer As Long, ByVal val As Long)
    Dim ally As Long
    If Not ULT(viewer, MAX_PLAYERS) Then Exit Sub

    'forward out vision to our allies
    For ally = 0 To MAX_PLAYERS - 1
        If hasSharedVision(viewer, ally) Then
            setSeenThisTick psObj, ally, MAX_(seenThisTick_(psObj, ally), val)
        End If
    Next
End Sub

Private Sub setSeenByInstantly(ByVal psObj As Long, ByVal viewer As Long, ByVal val As Long)
    Dim ally As Long
    If Not ULT(viewer, MAX_PLAYERS) Then Exit Sub

    'forward out vision to our allies
    For ally = 0 To MAX_PLAYERS - 1
        If hasSharedVision(viewer, ally) Then
            setSeenThisTick psObj, ally, MAX_(seenThisTick_(psObj, ally), val)
            BASE_OBJECT_visible(psObj, ally) = MAX_(BASE_OBJECT_visible(psObj, ally), val)
        End If
    Next
End Sub

' Calculate which objects we should know about based on alliances and satellite view.
Private Sub processVisibilitySelf(ByVal psObj As Long)
    Dim viewer As Long, psStruct As Long, psDroid As Long, objPlayer As Long, c As Boolean

    objPlayer = SIMPLE_OBJECT_player(psObj)
    If SIMPLE_OBJECT_type(psObj) <> OBJ_FEATURE Then
        If objSensorRange(psObj) > 0 Then
            ' one can trivially see oneself
            setSeenBy psObj, objPlayer, UBYTE_MAX
        End If
    End If

    ' if a player has a SAT_UPLINK structure, or has godMode enabled (and is not a spectator!!),
    ' they can see everything!
    For viewer = 0 To MAX_PLAYERS - 1
        c = getSatUplinkExists(viewer)
        If Not c Then
            If viewer = selectedPlayer And godMode Then c = Not NetPlay_players_isSpectator(viewer)
        End If
        If c Then
            setSeenBy psObj, viewer, UBYTE_MAX
        End If
    Next

    BASE_OBJECT_flags(psObj) = BASE_OBJECT_flags(psObj) And Not BIT_(OBJECT_FLAG_TARGETED) ' Remove any targetting locks from last update.

    ' If we're a CB sensor, make our target visible instantly. Although this is actually checking visibility of our target, we do it here anyway.
    psStruct = castStructure(psObj)
    ' you can always see anything that a CB sensor is targetting
    ' Anyone commenting this out again will get a knee capping from John.
    ' You have been warned!!
    If psStruct <> 0 Then
        If aStruct(psStruct).status = SS_BUILT Then
            If structCBSensor(psStruct) Or structVTOLCBSensor(psStruct) Then
                If aStruct(psStruct).psTarget(0) <> 0 Then
                    setSeenByInstantly aStruct(psStruct).psTarget(0), objPlayer, UBYTE_MAX
                End If
            End If
        End If
    End If
    psDroid = castDroid(psObj)
    If psDroid <> 0 Then
        If aDroid(psDroid).action = DACTION_OBSERVE Then
            If cbSensorDroid(psDroid) And aDroid(psDroid).psActionTarget(0) <> 0 Then
                ' Anyone commenting this out will get a knee capping from John.
                ' You have been warned!!
                setSeenByInstantly aDroid(psDroid).psActionTarget(0), objPlayer, UBYTE_MAX
            End If
        End If
    End If
End Sub

' Calculate which objects we can see. Better to call after processVisibilitySelf, since that check is cheaper.
Private Sub processVisibilityVision(ByVal psViewer As Long)
    Dim n As Long, gi As Long, psObj As Long, val As Long, pos As Vector3i, viewerPlayer As Long

    If SIMPLE_OBJECT_type(psViewer) = OBJ_FEATURE Then
        Exit Sub
    End If

    ' get all the objects from the grid the droid is in
    ' Will give inconsistent results if hasSharedVision is not an equivalence relation.
    pos = SIMPLE_OBJECT_pos(psViewer)
    viewerPlayer = SIMPLE_OBJECT_player(psViewer)
    n = gridStartIterateUnseen(pos.x, pos.y, objSensorRange(psViewer), viewerPlayer, gridListVision)
    For gi = 0 To n - 1
        psObj = gridListVision(gi)

        val = visibleObject(psViewer, psObj, False)

        ' If we've got ranged line of sight...
        If val > 0 Then
            ' Tell system that this side can see this object
            setSeenBy psObj, viewerPlayer, val

            ' Check if scripting system wants to trigger an event for this
            triggerEventSeen psViewer, psObj
        End If
    Next
End Sub

' Find out what can see this object
' Fade in/out of view. Must be called after calculation of which objects are seen.
Private Sub processVisibilityLevel(ByVal psObj As Long, addedMessage As Boolean)
    Dim player As Long, justBecameVisible As Boolean, visLevel As Long, vis As Long, psMessage As Long
    Dim typ As Long, objType As OBJECT_TYPE, pos As Vector3i

    objType = SIMPLE_OBJECT_type(psObj)
    ' update the visibility levels
    For player = 0 To MAX_PLAYERS - 1
        justBecameVisible = False
        visLevel = seenThisTick_(psObj, player)

        If player = SIMPLE_OBJECT_player(psObj) Then
            ' owner can always see it fully
            BASE_OBJECT_visible(psObj, player) = UBYTE_MAX
            GoTo continue_
        End If

        ' Droids can vanish from view, other objects will stay
        vis = BASE_OBJECT_visible(psObj, player)
        If objType <> OBJ_DROID Then
            visLevel = MAX_(visLevel, vis)
        End If

        If visLevel > vis Then
            justBecameVisible = vis <= 0

            BASE_OBJECT_visible(psObj, player) = MIN_(vis + visLevelInc, visLevel)
        ElseIf visLevel < vis Then
            BASE_OBJECT_visible(psObj, player) = MAX_(vis - visLevelDec, visLevel)
        End If

        If justBecameVisible Then
            ' Make sure all tiles under a feature/structure become visible when you see it
            If objType = OBJ_STRUCTURE Or objType = OBJ_FEATURE Then
                setUnderTilesVis psObj, player
            End If

            ' if a feature has just become visible set the message blips
            If objType = OBJ_FEATURE Then
                typ = NO_SOUND
                pos = aFeature(psObj).pos

                ' If this is an oil resource we want to add a proximity message for
                ' the selected Player - if there isn't an Resource Extractor on it.
                If asFeatureStats(aFeature(psObj).psStats).subType = FEAT_OIL_RESOURCE And Not TileHasStructure(mapTile(map_coord(pos.x), map_coord(pos.y))) Then
                    typ = ID_SOUND_RESOURCE_HERE
                ElseIf asFeatureStats(aFeature(psObj).psStats).subType = FEAT_GEN_ARTE Then
                    typ = ID_SOUND_ARTEFACT_DISC
                End If

                If typ <> NO_SOUND Then
                    psMessage = addMessage(MSG_PROXIMITY, True, player)
                    If psMessage <> 0 Then
                        MESSAGE_psObj(psMessage) = psObj
                        debug_ 821, LOG_MSG, "processVisibilityLevel", "Added message for oil well or artefact, pViewData=%p", MESSAGE_pViewData(psMessage)
                        addedMessage = True
                    End If
                    If Not bInTutorial And player = selectedPlayer Then
                        ' play message to indicate been seen
                        audio_QueueTrackPos typ, pos.x, pos.y, pos.z
                    End If
                End If
            End If
        End If
continue_:
    Next
End Sub

Public Sub processVisibility()
    Dim player As Long, it As Long, it2 As Long, psObj As Long, psTarget As Long, addedMessage As Boolean, c As Boolean
    Dim opos As Vector3i, tpos As Vector3i

    updateSpotters
    For player = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(player).head
        Do While it <> 0
            processVisibilitySelf lnObj(it)
            it = lnNext(it)
        Loop
        it = apsStructLists(player).head
        Do While it <> 0
            processVisibilitySelf lnObj(it)
            it = lnNext(it)
        Loop
    Next
    it = apsFeatureList(0).head
    Do While it <> 0
        processVisibilitySelf lnObj(it)
        it = lnNext(it)
    Loop
    For player = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(player).head
        Do While it <> 0
            processVisibilityVision lnObj(it)
            it = lnNext(it)
        Loop
        it = apsStructLists(player).head
        Do While it <> 0
            processVisibilityVision lnObj(it)
            it = lnNext(it)
        Loop
    Next
    it = apsSensorList(0).head
    Do While it <> 0
        psObj = lnObj(it)
        If objRadarDetector(psObj) Then
            it2 = apsSensorList(0).head
            Do While it2 <> 0
                psTarget = lnObj(it2)
                c = False
                If psObj <> psTarget Then
                    If BASE_OBJECT_visible(psTarget, SIMPLE_OBJECT_player(psObj)) < UBYTE_MAX \ 2 Then
                        If objActiveRadar(psTarget) <> -1 Then
                            opos = SIMPLE_OBJECT_pos(psObj)
                            tpos = SIMPLE_OBJECT_pos(psTarget)
                            c = iHypot(tpos.x - opos.x, tpos.y - opos.y) < objSensorRange(psObj) * 10
                        End If
                    End If
                End If
                If c Then
                    BASE_OBJECT_visible(psTarget, SIMPLE_OBJECT_player(psObj)) = UBYTE_MAX \ 2
                End If
                it2 = lnNext(it2)
            Loop
        End If
        it = lnNext(it)
    Loop
    addedMessage = False
    For player = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(player).head
        Do While it <> 0
            processVisibilityLevel lnObj(it), addedMessage
            it = lnNext(it)
        Loop
        it = apsStructLists(player).head
        Do While it <> 0
            processVisibilityLevel lnObj(it), addedMessage
            it = lnNext(it)
        Loop
    Next
    it = apsFeatureList(0).head
    Do While it <> 0
        processVisibilityLevel lnObj(it), addedMessage
        it = lnNext(it)
    Loop
    If addedMessage Then
        jsDebugMessageUpdate
    End If
End Sub

Public Sub setUnderTilesVis(ByVal psObj As Long, ByVal player As Long)
    Dim i As Long, j As Long, mapX As Long, mapY As Long, width As Long, breadth As Long, psTile As Long

    If Not ULT(player, MAX_PLAYERS) Then
        Exit Sub
    End If

    If SIMPLE_OBJECT_type(psObj) = OBJ_FEATURE Then
        width = asFeatureStats(aFeature(psObj).psStats).baseWidth
        breadth = asFeatureStats(aFeature(psObj).psStats).baseBreadth
        mapX = map_coord(aFeature(psObj).pos.x - width * TILE_UNITS \ 2)
        mapY = map_coord(aFeature(psObj).pos.y - breadth * TILE_UNITS \ 2)
    Else
        ' Must be a structure
        width = asStructureStats(aStruct(psObj).pStructureType).baseWidth
        breadth = asStructureStats(aStruct(psObj).pStructureType).baseBreadth
        mapX = map_coord(aStruct(psObj).pos.x - width * TILE_UNITS \ 2)
        mapY = map_coord(aStruct(psObj).pos.y - breadth * TILE_UNITS \ 2)
    End If

    For i = 0 To width   ' + 1 because visibility is for top left of tile.
        For j = 0 To breadth   ' + 1 because visibility is for top left of tile.
            psTile = mapTile(mapX + i, mapY + j)
            psMapTiles(psTile).tileExploredBits = psMapTiles(psTile).tileExploredBits Or alliancebits(player)
        Next
    Next
End Sub

' fast test for whether obj2 is in range of obj1
Public Function visObjInRange(ByVal psObj1 As Long, ByVal psObj2 As Long, ByVal range As Long) As Boolean
    Dim p1 As Vector3i, p2 As Vector3i, xdiff As Long, ydiff As Long
    p1 = SIMPLE_OBJECT_pos(psObj1)
    p2 = SIMPLE_OBJECT_pos(psObj2)
    xdiff = p1.x - p2.x
    ydiff = p1.y - p2.y
    If Abs(xdiff) <= range And Abs(ydiff) <= range Then visObjInRange = xdiff * xdiff + ydiff * ydiff <= range
End Function

' If we have ECM, use this for range instead. Otherwise, the sensor's range will be used for
' jamming range, which we do not want. Rather limit ECM unit sensor range to jammer range.
Public Function objSensorRange(ByVal psObj As Long) As Long
    Dim ecmRange_ As Long, player As Long
    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        player = aDroid(psObj).player
        ecmRange_ = asECMStats(aDroid(psObj).asBits(COMP_ECM)).upgrade(player).range
        If ecmRange_ > 0 Then
            objSensorRange = ecmRange_
            Exit Function
        End If
        objSensorRange = asSensorStats(aDroid(psObj).asBits(COMP_SENSOR)).upgrade(player).range
    Case OBJ_STRUCTURE
        player = aStruct(psObj).player
        ecmRange_ = asECMStats(asStructureStats(aStruct(psObj).pStructureType).pECM).upgrade(player).range
        If ecmRange_ <> 0 Then
            objSensorRange = ecmRange_
            Exit Function
        End If
        objSensorRange = asSensorStats(asStructureStats(aStruct(psObj).pStructureType).pSensor).upgrade(player).range
    Case Else
        objSensorRange = 0
    End Select
End Function

Public Function objJammerPower(ByVal psObj As Long) As Long
    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        objJammerPower = asECMStats(aDroid(psObj).asBits(COMP_ECM)).upgrade(aDroid(psObj).player).range
    Case OBJ_STRUCTURE
        objJammerPower = asECMStats(asStructureStats(aStruct(psObj).pStructureType).pECM).upgrade(aStruct(psObj).player).range
    Case Else
        objJammerPower = 0
    End Select
End Function
' ---- B7: the line of sight and the line of fire

' The los ray callback (RAY_CALLBACK_LOS; its data is losHelp)
Public Function rayLOSCallback(pos As Vector2i, ByVal dist As Long) As Boolean
    Dim newGrad As Long, tileX As Long, tileY As Long, psTile As Long, psStruct As Long

    If Not (pos.x >= 0 And pos.x < world_coord(mapWidth) And pos.y >= 0 And pos.y < world_coord(mapHeight)) Then ASSERT_FAILURE "pos.x >= 0 && pos.x < world_coord(mapWidth) && pos.y >= 0 && pos.y < world_coord(mapHeight)", "src/visibility.cpp:347", "rayLOSCallback", "rayLOSCallback: coords off map"

    If losHelp.rayStart Then
        losHelp.rayStart = False
    Else
        ' Calculate the current LOS gradient
        newGrad = (losHelp.lastHeight - losHelp.startHeight) * GRAD_MUL \ MAX_(1, losHelp.lastDist)
        If newGrad >= losHelp.currGrad Then
            losHelp.currGrad = newGrad
        End If
    End If

    losHelp.lastDist = dist
    losHelp.lastHeight = map_Height(pos.x, pos.y)

    If losHelp.wallsBlock Then
        ' Store the height at this tile for next time round
        tileX = map_coord(pos.x)
        tileY = map_coord(pos.y)

        If tileX <> losHelp.final.x Or tileY <> losHelp.final.y Then
            psTile = mapTile(tileX, tileY)
            If TileHasWall_raycast(psTile) And Not TileHasSmallStructure(psTile) Then
                psStruct = psMapTiles(psTile).psObject
                If asStructureStats(aStruct(psStruct).pStructureType).typ <> REF_GATE Or aStruct(psStruct).state <> SAS_OPEN Then
                    losHelp.lastHeight = 2 * TILE_MAX_HEIGHT
                    losHelp.wall = pos
                    losHelp.numWalls = losHelp.numWalls + 1
                End If
            End If
        End If
    End If

    rayLOSCallback = True
End Function

' Check whether psViewer can see psTarget.
' psViewer should be an object that has some form of sensor,
' currently droids and structures.
' psTarget can be any type of BASE_OBJECT (e.g. a tree).
' wallsBlock controls whether structures block LOS
Public Function visibleObject(ByVal psViewer As Long, ByVal psTarget As Long, ByVal wallsBlock As Boolean) As Long
    Dim range As Long, vPos As Vector3i, tPos As Vector3i, psStruct As Long, dist As Long, psTile As Long
    Dim jammed As Boolean, tileWatched As Boolean, tileWatchedSensor As Boolean, vPlayer As Long, cond As Boolean
    Dim src As Vector2i, dst As Vector2i

    If Not (psViewer <> 0) Then ASSERT_FAILURE "psViewer != nullptr", "src/visibility.cpp:487", "visibleObject", "Invalid viewer pointer!": visibleObject = 0: Exit Function
    If Not (psTarget <> 0) Then ASSERT_FAILURE "psTarget != nullptr", "src/visibility.cpp:488", "visibleObject", "Invalid viewed pointer!": visibleObject = 0: Exit Function

    range = objSensorRange(psViewer)
    vPos = SIMPLE_OBJECT_pos(psViewer)
    tPos = SIMPLE_OBJECT_pos(psTarget)

    If Not worldOnMap(vPos.x, vPos.y) Or Not worldOnMap(tPos.x, tPos.y) Then
        'Most likely a VTOL or transporter
        debug_ 495, LOG_WARNING, "visibleObject", "Trying to view something off map!"
        visibleObject = 0
        Exit Function
    End If

    ' Get the sensor range
    Select Case SIMPLE_OBJECT_type(psViewer)
    Case OBJ_DROID
        If aDroid(psViewer).order.psObj = psTarget Then
            If cbSensorDroid(psViewer) Then
                ' if it is targetted by a counter battery sensor, it is seen
                visibleObject = UBYTE_MAX
                Exit Function
            End If
        End If
    Case OBJ_STRUCTURE
        psStruct = psViewer

        ' a structure that is being built cannot see anything
        If aStruct(psStruct).status <> SS_BUILT Then
            visibleObject = 0
            Exit Function
        End If

        If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_WALL _
           Or asStructureStats(aStruct(psStruct).pStructureType).typ = REF_GATE _
           Or asStructureStats(aStruct(psStruct).pStructureType).typ = REF_WALLCORNER Then
            visibleObject = 0
            Exit Function
        End If

        cond = SIMPLE_OBJECT_type(psTarget) = OBJ_DROID
        If cond Then cond = DROID_isVtol(psTarget)
        If cond Then cond = asWeaponStats(BASE_OBJECT_getWeaponStats(psStruct, 0)).surfaceToAir = SHOOT_IN_AIR
        If cond Then
            range = 3 * range \ 2  ' increase vision range of AA vs VTOL
        End If

        If aStruct(psStruct).psTarget(0) = psTarget Then
            If structCBSensor(psStruct) Or structVTOLCBSensor(psStruct) Then
                ' if a unit is targetted by a counter battery sensor
                ' it is automatically seen
                visibleObject = UBYTE_MAX
                Exit Function
            End If
        End If
    Case Else
        ASSERT_FAILURE "false", "src/visibility.cpp:545", "visibleObject", "Visibility checking is only implemented for units and structures"
        visibleObject = 0
        Exit Function
    End Select

    ' First see if the target is in sensor range
    dist = iHypot(tPos.x - vPos.x, tPos.y - vPos.y)
    If dist = 0 Then
        visibleObject = UBYTE_MAX ' Should never be on top of each other, but ...
        Exit Function
    End If

    psTile = mapTile(map_coord(tPos.x), map_coord(tPos.y))
    vPlayer = SIMPLE_OBJECT_player(psViewer)
    jammed = (psMapTiles(psTile).jammerBits And Not alliancebits(vPlayer)) <> 0

    ' Special rule for VTOLs, as they are not affected by ECM
    cond = False
    If SIMPLE_OBJECT_type(psTarget) = OBJ_DROID Then cond = DROID_isVtol(psTarget)
    If Not cond And SIMPLE_OBJECT_type(psViewer) = OBJ_DROID Then cond = DROID_isVtol(psViewer)
    If cond And dist < range Then
        visibleObject = UBYTE_MAX
        Exit Function
    End If

    ' initialise the callback variables
    losHelp.rayStart = True
    losHelp.wallsBlock = wallsBlock
    losHelp.startHeight = vPos.z + map_Height(vPos.x, vPos.y)
    losHelp.final.x = map_coord(tPos.x)
    losHelp.final.y = map_coord(tPos.y)
    losHelp.lastHeight = 0
    losHelp.lastDist = 0
    losHelp.currGrad = -UBYTE_MAX * GRAD_MUL * ELEVATION_SCALE
    losHelp.numWalls = 0
    losHelp.wall.x = 0
    losHelp.wall.y = 0

    ' Cast a ray from the viewer to the target
    src.x = vPos.x: src.y = vPos.y
    dst.x = tPos.x: dst.y = tPos.y
    rayCast src, dst, RAY_CALLBACK_LOS

    If gWallSet Then ' Out globals are set
        gWall = losHelp.wall
        gNumWalls = losHelp.numWalls
    End If

    tileWatched = psMapTiles(psTile).watchers(vPlayer) > 0
    tileWatchedSensor = psMapTiles(psTile).sensors(vPlayer) > 0

    ' Show objects hidden by ECM jamming with radar blips
    If jammed Then
        If Not tileWatched And tileWatchedSensor Then
            visibleObject = UBYTE_MAX \ 2
            Exit Function
        End If
    End If
    ' Show objects that are seen directly
    If tileWatched Or (Not jammed And tileWatchedSensor) Then
        visibleObject = UBYTE_MAX
        Exit Function
    End If
    ' Show detected sensors as radar blips
    If objRadarDetector(psViewer) Then
        If objActiveRadar(psTarget) <> -1 Then
            If dist < range * 10 Then
                visibleObject = UBYTE_MAX \ 2
                Exit Function
            End If
        End If
    End If
    ' else not seen
    visibleObject = 0
End Function

' Find the wall that is blocking LOS to a target (if any)
Public Function visGetBlockingWall(ByVal psViewer As Long, ByVal psTarget As Long) As Long
    Dim numWalls As Long, wall_ As Vector2i, tileX As Long, tileY As Long, player As Long, it As Long, psWall As Long

    numWalls = 0

    ' HACK Using globals to not clutter visibleObject() interface too much
    gNumWalls = numWalls
    gWall = wall_
    gWallSet = True

    visibleObject psViewer, psTarget, True

    gWallSet = False
    numWalls = gNumWalls
    wall_ = gWall

    ' see if there was a wall in the way
    If numWalls > 0 Then
        tileX = map_coord(wall_.x)
        tileY = map_coord(wall_.y)

        For player = 0 To MAX_PLAYERS - 1
            it = apsStructLists(player).head
            Do While it <> 0
                psWall = lnObj(it)
                If map_coord(aStruct(psWall).pos.x) = tileX And map_coord(aStruct(psWall).pos.y) = tileY Then
                    visGetBlockingWall = psWall
                    Exit Function
                End If
                it = lnNext(it)
            Loop
        Next
    End If
    visGetBlockingWall = 0
End Function

' Check whether psViewer can fire directly at psTarget.
' psTarget can be any type of BASE_OBJECT (e.g. a tree).
Public Function lineOfFire(ByVal psViewer As Long, ByVal psTarget As Long, ByVal weapon_slot As Long, ByVal wallsBlock As Boolean) As Boolean
    Dim psStats As Long, distance As Long, range As Long, min_angle As Long, vPos As Vector3i, tPos As Vector3i

    psStats = -1

    If Not (psViewer <> 0) Then ASSERT_FAILURE "psViewer != nullptr", "src/visibility.cpp:959", "lineOfFire", "Invalid shooter pointer!": lineOfFire = False: Exit Function
    If Not (psTarget <> 0) Then ASSERT_FAILURE "psTarget != nullptr", "src/visibility.cpp:960", "lineOfFire", "Invalid target pointer!": lineOfFire = False: Exit Function
    If Not (SIMPLE_OBJECT_type(psViewer) = OBJ_DROID Or SIMPLE_OBJECT_type(psViewer) = OBJ_STRUCTURE) Then ASSERT_FAILURE "psViewer->type == OBJ_DROID || psViewer->type == OBJ_STRUCTURE", "src/visibility.cpp:961", "lineOfFire", "Bad viewer type": lineOfFire = False: Exit Function

    If SIMPLE_OBJECT_type(psViewer) = OBJ_DROID Then
        psStats = DROID_getWeaponStats(psViewer, weapon_slot)
    ElseIf SIMPLE_OBJECT_type(psViewer) = OBJ_STRUCTURE Then
        psStats = BASE_OBJECT_getWeaponStats(psViewer, weapon_slot)
    End If
    ' 2d distance
    vPos = SIMPLE_OBJECT_pos(psViewer)
    tPos = SIMPLE_OBJECT_pos(psTarget)
    distance = iHypot(tPos.x - vPos.x, tPos.y - vPos.y)
    range = proj_GetLongRange(asWeaponStats(psStats), SIMPLE_OBJECT_player(psViewer))
    If proj_Direct(psStats) Then
        ' direct shots could collide with ground
        lineOfFire = False
        If range >= distance Then
            lineOfFire = LINE_OF_FIRE_MINIMUM <= checkFireLine(psViewer, psTarget, weapon_slot, wallsBlock, True)
        End If
    Else
        ' indirect shots always have a line of fire, IF the forced
        ' minimum angle doesn't move it out of range
        min_angle = checkFireLine(psViewer, psTarget, weapon_slot, wallsBlock, False)
        ' NOTE This code seems similar to the code in combFire in combat.cpp.
        If min_angle > DEG(PROJ_MAX_PITCH) Then
            If iSin(2 * min_angle) < iSin(2 * DEG(PROJ_MAX_PITCH)) Then
                range = MUL32(range, iSin(2 * min_angle)) \ iSin(2 * DEG(PROJ_MAX_PITCH))
            End If
        End If
        lineOfFire = range >= distance
    End If
End Function

' Check how much of psTarget is hitable from psViewer's gun position
Public Function areaOfFire(ByVal psViewer As Long, ByVal psTarget As Long, ByVal weapon_slot As Long, ByVal wallsBlock As Boolean) As Long
    If psViewer = 0 Then
        areaOfFire = 0  ' Lassat special case, avoid assertion.
        Exit Function
    End If

    areaOfFire = checkFireLine(psViewer, psTarget, weapon_slot, wallsBlock, True)
End Function

' Check the minimum angle to hitpsTarget from psViewer via indirect shots
Public Function arcOfFire(ByVal psViewer As Long, ByVal psTarget As Long, ByVal weapon_slot As Long, ByVal wallsBlock As Boolean) As Long
    arcOfFire = checkFireLine(psViewer, psTarget, weapon_slot, wallsBlock, False)
End Function

' helper function for checkFireLine (angletan: an int64_t, a Double)
Private Sub angle_check(angletan As Double, ByVal positionSq As Long, ByVal height As Long, ByVal distanceSq As Long, ByVal targetHeight As Long, ByVal direct As Boolean)
    Dim current As Double, dist As Long, pos As Long, num As Double, den As Double
    If direct Then
        current = (65536 * height) \ iSqrt(positionSq)
    Else
        dist = iSqrt(distanceSq)
        pos = iSqrt(positionSq)
        current = (pos * targetHeight) \ dist
        If current < height And pos > TILE_UNITS \ 2 And pos < dist - TILE_UNITS \ 2 Then
            ' solve the following trajectory parabolic equation
            ' ( targetHeight ) = a * distance^2 + factor * distance
            ' ( height ) = a * position^2 + factor * position
            '  "a" depends on angle, gravity and shooting speed.
            '  luckily we don't need it for this at all, since
            ' factor = tan(firing_angle)
            num = CDbl(distanceSq) * height - CDbl(positionSq) * targetHeight
            den = CDbl(distanceSq) * pos - CDbl(dist) * positionSq
            If Abs(num) < 137438953472# Then    ' 2^37: 65536 * num < 2^53, so Fix of the Double quotient is C's
                current = Fix(65536# * num / den)
            Else
                current = CDbl(Fix(CDec(num) * 65536 / CDec(den)))
            End If
        Else
            current = 0
        End If
    End If
    If current > angletan Then angletan = current   ' std::max
End Sub

' Check fire line from psViewer to psTarget
' psTarget can be any type of BASE_OBJECT (e.g. a tree).
Private Function checkFireLine(ByVal psViewer As Long, ByVal psTarget As Long, ByVal weapon_slot As Long, ByVal wallsBlock As Boolean, ByVal direct As Boolean) As Long
    Dim pos As Vector3i, dest As Vector3i, start As Vector2i, diff As Vector2i, current As Vector2i, halfway As Vector2i
    Dim nxt As Vector2i, part As Vector2i, muzzle As Vector3i, distSq As Long, partSq As Long, oldPartSq As Long
    Dim angletan As Double, hasSplitIntersection As Boolean, psTile As Long, tileObj As Long, a As Long

    If Not (psViewer <> 0) Then ASSERT_FAILURE "psViewer != nullptr", "src/visibility.cpp:1059", "checkFireLine", "Invalid shooter pointer!"
    If Not (psTarget <> 0) Then ASSERT_FAILURE "psTarget != nullptr", "src/visibility.cpp:1060", "checkFireLine", "Invalid target pointer!"
    If psViewer = 0 Or psTarget = 0 Then
        checkFireLine = -1
        Exit Function
    End If

    ' CorvusCorax: get muzzle offset (code from projectile.c)
    If SIMPLE_OBJECT_type(psViewer) = OBJ_DROID And weapon_slot >= 0 Then
        calcDroidMuzzleBaseLocation psViewer, muzzle, weapon_slot
    ElseIf SIMPLE_OBJECT_type(psViewer) = OBJ_STRUCTURE And weapon_slot >= 0 Then
        calcStructureMuzzleBaseLocation psViewer, muzzle, weapon_slot
    Else ' incase anything wants a projectile
        muzzle = SIMPLE_OBJECT_pos(psViewer)
    End If

    pos = muzzle
    dest = SIMPLE_OBJECT_pos(psTarget)
    diff.x = dest.x - pos.x
    diff.y = dest.y - pos.y

    distSq = dot2i(diff, diff)
    If distSq = 0 Then
        ' Should never be on top of each other, but ...
        checkFireLine = 1000
        Exit Function
    End If

    current.x = pos.x
    current.y = pos.y
    start = current
    angletan = -1000# * 65536#
    partSq = 0
    ' run a manual trace along the line of fire until target is reached
    Do While partSq < distSq
        oldPartSq = partSq

        If partSq > 0 Then
            angle_check angletan, partSq, map_Height(current.x, current.y) - pos.z, distSq, dest.z - pos.z, direct
        End If

        ' intersect current tile with line of fire
        nxt = diff
        hasSplitIntersection = map_Intersect(current.x, current.y, nxt.x, nxt.y, halfway.x, halfway.y)

        If hasSplitIntersection Then
            ' check whether target was reached before tile split line:
            part.x = halfway.x - start.x
            part.y = halfway.y - start.y
            partSq = dot2i(part, part)

            If partSq >= distSq Then
                Exit Do
            End If

            If partSq > 0 Then
                angle_check angletan, partSq, map_Height(halfway.x, halfway.y) - pos.z, distSq, dest.z - pos.z, direct
            End If
        End If

        ' check for walls and other structures
        ' TODO: if there is a structure on the same tile as the shooter (and the shooter is not that structure) check if LOF is blocked by it.
        If wallsBlock And oldPartSq > 0 Then
            halfway.x = current.x + (nxt.x - current.x) \ 2
            halfway.y = current.y + (nxt.y - current.y) \ 2
            psTile = mapTile(map_coord(halfway.x), map_coord(halfway.y))
            If TileHasStructure(psTile) And psMapTiles(psTile).psObject <> psTarget Then
                ' check whether target was reached before tile's "half way" line
                part.x = halfway.x - start.x
                part.y = halfway.y - start.y
                partSq = dot2i(part, part)

                If partSq >= distSq Then
                    Exit Do
                End If

                ' allowed to shoot over enemy structures if they are NOT the target
                If partSq > 0 Then
                    tileObj = psMapTiles(psTile).psObject
                    angle_check angletan, oldPartSq, _
                                SIMPLE_OBJECT_pos(tileObj).z + establishTargetHeight(tileObj) - pos.z, _
                                distSq, dest.z - pos.z, direct
                End If
            End If
        End If

        ' next
        current = nxt
        part.x = current.x - start.x
        part.y = current.y - start.y
        partSq = dot2i(part, part)
        If Not (partSq > oldPartSq) Then ASSERT_FAILURE "partSq > oldPartSq", "src/visibility.cpp:1160", "checkFireLine", "areaOfFire(): no progress in tile-walk! From: %i,%i to %i,%i stuck in %i,%i", map_coord(pos.x), map_coord(pos.y), map_coord(dest.x), map_coord(dest.y), map_coord(current.x), map_coord(current.y)
    Loop
    If direct Then
        checkFireLine = FromU32(establishTargetHeight(psTarget) - (pos.z + Fix(angletan * iSqrt(distSq) / 65536#) - dest.z))
    Else
        a = iAtan2(FromU32(angletan), 65536)
        a = angleDelta(a)
        checkFireLine = DEG(1) + a
    End If
End Function
