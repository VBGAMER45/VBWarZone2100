Attribute VB_Name = "fpath_c"
' src/fpath.cpp, fpath.h: the interface to the routing functions: whether a
' map tile blocks a propulsion type, the path jobs and their results, and
' fpathDroidRoute. fpathTest (a unit test of the threads) is not ported.
' (VB) There are no path-finding threads: a job runs (fpathExecute, on
' astar_c's one context list) when fpathRoute queues it, and its result
' waits in pathResults for the droid's next call, as the future does. The
' C++ sends every job that could reuse a PathfindContext of another to the
' same thread, in queue order (fpathJobDispatchThreadId), and the job's
' blocking map is fixed when it is queued, so the paths do not depend on
' when or where a job runs. "Using threads: %zu" logs the number the C++
' would start. std::unordered_map<uint32_t, wz::future<PATHRESULT>>
' pathResults is a Scripting.Dictionary from the droid id to a slot of
' pathResultSlots.
Option Explicit

' Used to modify the pathfinding algorithm (enum FPATH_MOVETYPE)
Public Const FMT_MOVE As Long = 0       ' Move around all obstacles
Public Const FMT_ATTACK As Long = 1     ' Assume that we will destroy enemy obstacles
Public Const FMT_BLOCK As Long = 2      ' Don't go through obstacles, not even gates.

Public Type PATHJOB
    propulsion As Long                  ' PROPULSION_TYPE
    droidType As Long                   ' DROID_TYPE
    destX As Long
    destY As Long
    origX As Long
    origY As Long
    dstStructure As StructureBounds
    droidID As Long                     ' UDWORD
    moveType As Long                    ' FPATH_MOVETYPE
    owner As Long                       ' Player owner
    blockingMap As Long                 ' std::shared_ptr<const PathBlockingMap>: Map of blocking tiles (astar_c's handle).
    acceptNearest As Boolean
    deleted As Boolean                  ' Droid was deleted, so throw away result when complete. Must still process this PATHJOB, since processing order can affect resulting paths (but can't affect the path length).
End Type

' Return values for routing
Public Enum FPATH_RETVAL
    FPR_OK                              ' found a route
    FPR_FAILED                          ' failed to find a route
    FPR_WAIT                            ' route is being calculated by the path-finding thread
End Enum

Private Type PATHRESULT
    droidID As Long                     ' Unique droid ID.
    sMove As MOVE_CONTROL               ' New movement values for the droid.
    retval As FPATH_RETVAL              ' Result value from path-finding.
    originalDest As Vector2i            ' Used to check if the pathfinding job is to the right destination.
End Type

' If the path finding system is shutdown or not
Private fpathQuit As Boolean

Private fpathThreads As Long            ' (VB) fpathThreads.size(): the threads the C++ would run
Private pathResults As Object           ' Scripting.Dictionary: droid id -> slot of pathResultSlots
Private pathResultSlots() As PATHRESULT
Private pathResultSlots_size As Long

Private Const MAX_FPATH_THREADS As Long = 2

Private Function fpathDetermineNumberOfThreads() As Long
    Dim logicalCPUCount As Long
    logicalCPUCount = wzGetLogicalCPUCount()
    If logicalCPUCount <= 1 Then
        fpathDetermineNumberOfThreads = 1
        Exit Function
    End If
    ' subtract one for the main thread
    fpathDetermineNumberOfThreads = logicalCPUCount - 1
    If fpathDetermineNumberOfThreads > MAX_FPATH_THREADS Then fpathDetermineNumberOfThreads = MAX_FPATH_THREADS
End Function

' initialise the findpath module
Public Function fpathInitialise() As Boolean
    Dim numThreads As Long
    ' The path system is up
    fpathQuit = False

    If fpathThreads = 0 Then
        numThreads = fpathDetermineNumberOfThreads()
        debug_ 166, LOG_INFO, "fpathInitialise", "Using threads: %zu", numThreads
        fpathThreads = numThreads           ' (VB) no thread is started: the jobs run when queued
    End If

    fpathInitialise = True
End Function

Public Sub fpathShutdown()
    If fpathThreads <> 0 Then
        ' Signal the path finding thread(s) to quit
        fpathQuit = True
        fpathThreads = 0
        fpathContextsRelease                ' (VB) the threads' contexts end with them
    End If
    fpathHardTableReset
End Sub

' Updates the pathfinding system.
Public Sub fpathUpdate()
    ' Nothing now
End Sub

Private Function fpathPropulsionDomain(ByVal propulsion As Long) As Long
    Select Case propulsion
    Case PROPULSION_TYPE_LIFT: fpathPropulsionDomain = 1        ' Air
    Case PROPULSION_TYPE_PROPELLOR: fpathPropulsionDomain = 2   ' Water
    Case PROPULSION_TYPE_HOVER: fpathPropulsionDomain = 3       ' Land and water
    Case Else: fpathPropulsionDomain = 0                        ' Land
    End Select
End Function

' Returns true iff the parameters have equivalent behaviour in fpathBaseBlockingTile.
Public Function fpathIsEquivalentBlocking(ByVal propulsion1 As Long, ByVal player1 As Long, ByVal moveType1 As Long, _
                                          ByVal propulsion2 As Long, ByVal player2 As Long, ByVal moveType2 As Long) As Boolean
    Dim domain1 As Long, domain2 As Long
    domain1 = fpathPropulsionDomain(propulsion1)
    domain2 = fpathPropulsionDomain(propulsion2)

    If domain1 <> domain2 Then
        fpathIsEquivalentBlocking = False
        Exit Function
    End If

    If domain1 = fpathPropulsionDomain(PROPULSION_TYPE_LIFT) Then
        fpathIsEquivalentBlocking = True    ' Air units ignore move type and player.
        Exit Function
    End If

    If moveType1 <> moveType2 Or player1 <> player2 Then
        fpathIsEquivalentBlocking = False
        Exit Function
    End If

    fpathIsEquivalentBlocking = True
End Function

Private Function prop2bits(ByVal propulsion As Long) As Long
    Dim bits As Long

    Select Case propulsion
    Case PROPULSION_TYPE_LIFT
        bits = AIR_BLOCKED
    Case PROPULSION_TYPE_HOVER
        bits = FEATURE_BLOCKED
    Case PROPULSION_TYPE_PROPELLOR
        bits = FEATURE_BLOCKED Or LAND_BLOCKED
    Case Else
        bits = FEATURE_BLOCKED Or WATER_BLOCKED
    End Select
    prop2bits = bits
End Function

' Check if the map tile at a location blocks a droid
Public Function fpathBaseBlockingTile(ByVal x As Long, ByVal y As Long, ByVal propulsion As Long, ByVal mapIndex As Long, ByVal moveType As Long) As Boolean
    Dim aux As Long, auxMask As Long, unitbits As Long

    ' All tiles outside of the map and on map border are blocking.
    If x < 1 Or y < 1 Or x > mapWidth - 1 Or y > mapHeight - 1 Then
        fpathBaseBlockingTile = True
        Exit Function
    End If

    ' Check scroll limits (used in campaign to partition the map.
    If propulsion <> PROPULSION_TYPE_LIFT And (x < scrollMinX + 1 Or y < scrollMinY + 1 Or x >= scrollMaxX - 1 Or y >= scrollMaxY - 1) Then
        ' coords off map - auto blocking tile
        fpathBaseBlockingTile = True
        Exit Function
    End If
    aux = auxTile(x, y, mapIndex)

    auxMask = 0
    Select Case moveType
    Case FMT_MOVE: auxMask = AUXBITS_NONPASSABLE    ' do not wish to shoot our way through enemy buildings, but want to go through friendly gates (without shooting them)
    Case FMT_ATTACK: auxMask = AUXBITS_OUR_BUILDING ' move blocked by friendly building, assuming we do not want to shoot it up en route
    Case FMT_BLOCK: auxMask = AUXBITS_BLOCKING      ' Do not wish to tunnel through closed gates or buildings.
    End Select

    unitbits = prop2bits(propulsion)    ' TODO - cache prop2bits to psDroid, and pass in instead of propulsion type
    If (unitbits And FEATURE_BLOCKED) <> 0 And (aux And auxMask) <> 0 Then
        fpathBaseBlockingTile = True    ' move blocked by building, and we cannot or do not want to shoot our way through anything
        Exit Function
    End If

    ' the MAX hack below is because blockTile() range does not include player-specific versions...
    If mapIndex - MAX_PLAYERS > 0 Then
        fpathBaseBlockingTile = (blockTile(x, y, mapIndex - MAX_PLAYERS) And unitbits) <> 0   ' finally check if move is blocked by propulsion related factors
    Else
        fpathBaseBlockingTile = (blockTile(x, y, 0) And unitbits) <> 0
    End If
End Function

Public Function fpathDroidBlockingTile(ByVal psDroid As Long, ByVal x As Long, ByVal y As Long, ByVal moveType As Long) As Boolean
    fpathDroidBlockingTile = fpathBaseBlockingTile(x, y, asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType, aDroid(psDroid).player, moveType)
End Function

' Check if the map tile at a location blocks a droid
Public Function fpathBlockingTile(ByVal x As Long, ByVal y As Long, ByVal propulsion As Long) As Boolean
    fpathBlockingTile = fpathBaseBlockingTile(x, y, propulsion, 0, FMT_BLOCK)   ' with FMT_BLOCK, it is irrelevant which player is passed in
End Function

' Returns the closest non-blocking tile to pos, or returns pos if no non-blocking tiles are present within a 2 tile distance.
' (player: the C++ default is 0, moveType: FMT_BLOCK)
Private Function findNonblockingPosition(pos As Vector3i, ByVal propulsion As Long, Optional ByVal player As Long = 0, Optional ByVal moveType As Long = FMT_BLOCK) As Vector3i
    Dim centreTile As Vector2i, bestTile As Vector2i, bestDistSq As Long, x As Long, y As Long
    Dim tile As Vector2i, diff As Vector2i, distSq As Long, minCoord As Vector2i, maxCoord As Vector2i
    centreTile.x = map_coord(pos.x)
    centreTile.y = map_coord(pos.y)
    If Not fpathBaseBlockingTile(centreTile.x, centreTile.y, propulsion, player, moveType) Then
        findNonblockingPosition = pos       ' Fast case, pos is not on a blocking tile.
        Exit Function
    End If

    bestTile = centreTile
    bestDistSq = INT32_MAX
    For y = -2 To 2
        For x = -2 To 2
            tile.x = centreTile.x + x
            tile.y = centreTile.y + y
            diff.x = world_coord(tile.x) + TILE_UNITS \ 2 - pos.x
            diff.y = world_coord(tile.y) + TILE_UNITS \ 2 - pos.y
            distSq = dot2i(diff, diff)
            If distSq < bestDistSq Then
                If Not fpathBaseBlockingTile(tile.x, tile.y, propulsion, player, moveType) Then
                    bestTile = tile
                    bestDistSq = distSq
                End If
            End If
        Next
    Next

    ' Return point on tile closest to the original pos.
    minCoord.x = world_coord(bestTile.x)
    minCoord.y = world_coord(bestTile.y)
    maxCoord.x = minCoord.x + TILE_UNITS - 1
    maxCoord.y = minCoord.y + TILE_UNITS - 1

    findNonblockingPosition.x = MIN_(MAX_(pos.x, minCoord.x), maxCoord.x)
    findNonblockingPosition.y = MIN_(MAX_(pos.y, minCoord.y), maxCoord.y)
    findNonblockingPosition.z = pos.z
End Function

Private Sub fpathSetMove(psMoveCntl As MOVE_CONTROL, ByVal targetX As Long, ByVal targetY As Long)
    ReDim psMoveCntl.asPath(0 To 0)
    psMoveCntl.asPath_size = 1
    psMoveCntl.destination = Vector2i_(targetX, targetY)
    psMoveCntl.asPath(0) = Vector2i_(targetX, targetY)
End Sub

' Set a direct path to position.
' Plan a path from psDroid's current position to given position without
' taking obstructions into consideration.
' Used for instance by VTOLs. Function is thread-safe.
Public Sub fpathSetDirectRoute(ByVal psDroid As Long, ByVal targetX As Long, ByVal targetY As Long)
    fpathSetMove aDroid(psDroid).sMove, targetX, targetY
End Sub

' Clean up path jobs and results for a droid. Function is thread-safe.
Public Sub fpathRemoveDroidData(ByVal id As Long)
    If pathResults Is Nothing Then Set pathResults = CreateObject("Scripting.Dictionary")
    If pathResults.Exists(id) Then pathResults.Remove id
End Sub

' (VB) crcSumVector2i(crc, data, dataLen) over psMove.asPath (crc_c has no Vector2i)
Private Function crcSumVector2i(ByVal crc As Long, data() As Vector2i, ByVal dataLen As Long) As Long
    Dim i As Long
    For i = 0 To dataLen - 1
        crc = crcSumI32(crc, data(i).x)
        crc = crcSumI32(crc, data(i).y)
    Next
    crcSumVector2i = crc
End Function

' (VB) pathResults[id] = the job's result: a free slot, or a new one
Private Sub storePathResult(ByVal id As Long, result As PATHRESULT)
    Dim slot As Long, used() As Boolean, k As Variant
    If pathResultSlots_size > pathResults.count Then
        ReDim used(0 To pathResultSlots_size - 1)
        For Each k In pathResults.Keys
            used(pathResults(k)) = True
        Next
        For slot = 0 To pathResultSlots_size - 1
            If Not used(slot) Then Exit For
        Next
    Else
        slot = pathResultSlots_size
        pathResultSlots_size = pathResultSlots_size + 1
        ReDim Preserve pathResultSlots(0 To pathResultSlots_size - 1)
    End If
    pathResultSlots(slot) = result
    pathResults(id) = slot
End Sub

Private Function fpathRoute(psMove As MOVE_CONTROL, ByVal id As Long, ByVal startX As Long, ByVal startY As Long, ByVal tX As Long, ByVal tY As Long, _
                            ByVal propulsionType As Long, ByVal droidType As Long, ByVal moveType As Long, ByVal owner As Long, _
                            ByVal acceptNearest As Boolean, dstStructure As StructureBounds) As FPATH_RETVAL
    Dim result As PATHRESULT, correctDestination As Boolean, retval As FPATH_RETVAL, job As PATHJOB, isFirstJob As Boolean
    objTrace id, "fpathRoute", "called(*,id=%d,sx=%d,sy=%d,ex=%d,ey=%d,prop=%d,type=%d,move=%d,owner=%d)", id, startX, startY, tX, tY, propulsionType, droidType, moveType, owner

    If Not worldOnMap(startX, startY) Or Not worldOnMap(tX, tY) Then
        debug_ 447, LOG_ERROR, "fpathRoute", "Droid trying to find path to/from invalid location (%d %d) -> (%d %d).", startX, startY, tX, tY
        objTrace id, "fpathRoute", "Invalid start/end."
        syncDebug "fpathRoute", "fpathRoute(..., %d, %d, %d, %d, %d, %d, %d, %d, %d) = FPR_FAILED", id, startX, startY, tX, tY, propulsionType, droidType, moveType, owner
        fpathRoute = FPR_FAILED
        Exit Function
    End If

    ' don't have to do anything if already there
    If startX = tX And startY = tY Then
        ' return failed to stop them moving anywhere
        objTrace id, "fpathRoute", "Tried to move nowhere"
        syncDebug "fpathRoute", "fpathRoute(..., %d, %d, %d, %d, %d, %d, %d, %d, %d) = FPR_FAILED", id, startX, startY, tX, tY, propulsionType, droidType, moveType, owner
        fpathRoute = FPR_FAILED
        Exit Function
    End If

    ' Check if waiting for a result
    If pathResults Is Nothing Then Set pathResults = CreateObject("Scripting.Dictionary")
    Do While psMove.Status = MOVEWAITROUTE
        objTrace id, "fpathRoute", "Checking if we have a path yet"

        If Not (pathResults.Exists(id)) Then ASSERT_FAILURE "I != pathResults.end()", "src/fpath.cpp:468", "fpathRoute", "Missing path result promise": fpathRoute = FPR_FAILED: Exit Function
        result = pathResultSlots(pathResults(id))
        If Not (result.retval <> FPR_OK Or result.sMove.asPath_size > 0) Then ASSERT_FAILURE "result.retval != FPR_OK || result.sMove.asPath.size() > 0", "src/fpath.cpp:470", "fpathRoute", "Ok result but no path in list"

        ' Copy over select fields - preserve others
        psMove.destination = result.sMove.destination
        correctDestination = tX = result.originalDest.x And tY = result.originalDest.y
        psMove.pathIndex = 0
        psMove.Status = MOVENAVIGATE
        psMove.asPath = result.sMove.asPath
        psMove.asPath_size = result.sMove.asPath_size
        retval = result.retval
        If Not (retval <> FPR_OK Or psMove.asPath_size > 0) Then ASSERT_FAILURE "retval != FPR_OK || psMove->asPath.size() > 0", "src/fpath.cpp:479", "fpathRoute", "Ok result but no path after copy"

        ' Remove it from the result list
        pathResults.Remove id

        objTrace id, "fpathRoute", "Got a path to (%d, %d)! Length=%d Retval=%d", psMove.destination.x, psMove.destination.y, psMove.asPath_size, retval
        syncDebug "fpathRoute", "fpathRoute(..., %d, %d, %d, %d, %d, %d, %d, %d, %d) = %d, path[%d] = %08X->(%d, %d)", id, startX, startY, tX, tY, propulsionType, droidType, _
            moveType, owner, retval, psMove.asPath_size, Not crcSumVector2i(0, psMove.asPath, psMove.asPath_size), psMove.destination.x, psMove.destination.y

        If Not correctDestination Then
            GoTo queuePathfinding           ' Seems we got the result of an old pathfinding job for this droid, so need to pathfind again.
        End If

        fpathRoute = retval
        Exit Function
    Loop
queuePathfinding:

    ' We were not waiting for a result, and found no trivial path, so create new job and start waiting
    job.origX = startX
    job.origY = startY
    job.droidID = id
    job.destX = tX
    job.destY = tY
    job.dstStructure = dstStructure
    job.droidType = droidType
    job.propulsion = propulsionType
    job.moveType = moveType
    job.owner = owner
    job.acceptNearest = acceptNearest
    job.deleted = False
    fpathSetBlockingMap job

    debug_ 512, LOG_NEVER, "fpathRoute", "starting new job for droid %d 0x%x", id, id
    ' Clear any results or jobs waiting already. It is a vital assumption that there is only one
    ' job or result for each droid in the system at any time.
    fpathRemoveDroidData id

    ' (VB) the packaged task runs now, on the main thread, and its result waits as the future does
    isFirstJob = True
    result = fpathExecute(job)
    storePathResult id, result

    objTrace id, "fpathRoute", "Queued up a path-finding request to (%d, %d), at least %d items earlier in queue", tX, tY, -CLng(isFirstJob)
    syncDebug "fpathRoute", "fpathRoute(..., %d, %d, %d, %d, %d, %d, %d, %d, %d) = FPR_WAIT", id, startX, startY, tX, tY, propulsionType, droidType, moveType, owner
    fpathRoute = FPR_WAIT                   ' wait while polling result queue
End Function

' Find a route for an DROID to a location in world coordinates
Public Function fpathDroidRoute(ByVal psDroid As Long, ByVal tX As Long, ByVal tY As Long, ByVal moveType As Long) As FPATH_RETVAL
    Dim acceptNearest As Boolean, psPropStats As Long, startPos As Vector3i, endPos As Vector3i, dstStructure As StructureBounds
    Dim droidPropulsionType As Long
    psPropStats = DROID_getPropulsionStats(psDroid)

    ' override for AI to blast our way through stuff
    If Not isHumanPlayer(aDroid(psDroid).player) And moveType = FMT_MOVE Then
        If aDroid(psDroid).asWeaps(0).nStat = 0 Then moveType = FMT_MOVE Else moveType = FMT_ATTACK
    End If

    If Not (psPropStats <> -1) Then ASSERT_FAILURE "psPropStats != nullptr", "src/fpath.cpp:554", "fpathDroidRoute", "invalid propulsion stats pointer": fpathDroidRoute = FPR_FAILED: Exit Function
    If Not (aDroid(psDroid).typ = OBJ_DROID) Then ASSERT_FAILURE "psDroid->type == OBJ_DROID", "src/fpath.cpp:555", "fpathDroidRoute", "We got passed an object that isn't a DROID!": fpathDroidRoute = FPR_FAILED: Exit Function

    ' Check whether the start and end points of the route are blocking tiles and find an alternative if they are.
    startPos = aDroid(psDroid).pos
    endPos = Vector3i_(tX, tY, 0)
    dstStructure = getStructureBounds(psMapTiles(worldTile(endPos.x, endPos.y)).psObject)
    droidPropulsionType = asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType
    startPos = findNonblockingPosition(startPos, droidPropulsionType, aDroid(psDroid).player, moveType)
    If Not StructureBounds_valid(dstStructure) Then  ' If there's a structure over the destination, ignore it, otherwise pathfind from somewhere around the obstruction.
        endPos = findNonblockingPosition(endPos, droidPropulsionType, aDroid(psDroid).player, moveType)
    End If
    objTrace aDroid(psDroid).id, "fpathDroidRoute", "Want to go to (%d, %d) -> (%d, %d), going (%d, %d) -> (%d, %d)", map_coord(aDroid(psDroid).pos.x), map_coord(aDroid(psDroid).pos.y), _
        map_coord(tX), map_coord(tY), map_coord(startPos.x), map_coord(startPos.y), map_coord(endPos.x), map_coord(endPos.y)
    Select Case aDroid(psDroid).order.typ
    Case DORDER_BUILD, DORDER_LINEBUILD, DORDER_HELPBUILD, DORDER_DEMOLISH, DORDER_REPAIR
        If aDroid(psDroid).order.typ = DORDER_BUILD Or aDroid(psDroid).order.typ = DORDER_LINEBUILD Then  ' build a number of structures in a row (walls + bridges)
            dstStructure = getStructureBounds_structStats(aDroid(psDroid).order.psStats, aDroid(psDroid).order.pos, aDroid(psDroid).order.direction)  ' Just need to get close enough to build (can be diagonally), do not need to reach the destination tile.
        End If
        ' fallthrough: help to build a structure, demolish a structure, repair
        acceptNearest = False
    Case Else
        acceptNearest = True
    End Select
    fpathDroidRoute = fpathRoute(aDroid(psDroid).sMove, aDroid(psDroid).id, startPos.x, startPos.y, endPos.x, endPos.y, asPropulsionStats(psPropStats).propulsionType, _
                                 aDroid(psDroid).droidType, moveType, aDroid(psDroid).player, acceptNearest, dstStructure)
End Function

' Run only from path thread
Private Function fpathExecute(job As PATHJOB) As PATHRESULT
    Dim result As PATHRESULT, retval As ASR_RETVAL
    result.droidID = job.droidID
    result.retval = FPR_FAILED
    result.originalDest = Vector2i_(job.destX, job.destY)

    retval = fpathAStarRoute(result.sMove, job)

    If Not (retval <> ASR_OK Or result.sMove.asPath_size > 0) Then ASSERT_FAILURE "retval != ASR_OK || result.sMove.asPath.size() > 0", "src/fpath.cpp:597", "fpathExecute", "Ok result but no path in result"
    Select Case retval
    Case ASR_NEAREST
        If job.acceptNearest Then
            objTrace job.droidID, "fpathExecute", "** Nearest route -- accepted **"
            result.retval = FPR_OK
        Else
            objTrace job.droidID, "fpathExecute", "** Nearest route -- rejected **"
            result.retval = FPR_FAILED
        End If
    Case ASR_FAILED
        objTrace job.droidID, "fpathExecute", "** Failed route **"
        ' Is this really a good idea? Was in original code.
        If job.propulsion = PROPULSION_TYPE_LIFT And (job.droidType <> DROID_TRANSPORTER And job.droidType <> DROID_SUPERTRANSPORTER) Then
            objTrace job.droidID, "fpathExecute", "Doing fallback for non-transport VTOL"
            fpathSetMove result.sMove, job.destX, job.destY
            result.retval = FPR_OK
        Else
            result.retval = FPR_FAILED
        End If
    Case ASR_OK
        objTrace job.droidID, "fpathExecute", "Got route of length %d", result.sMove.asPath_size
        result.retval = FPR_OK
    End Select
    fpathExecute = result
End Function

' Quick O(1) test of whether it is theoretically possible to go from origin to destination
' using the given propulsion type. orig and dest are in world coordinates.
Public Function fpathCheck(orig As Vector3i, dest As Vector3i, ByVal propulsion As Long) As Boolean
    Dim origTile As Long, destTile As Long, p As Vector3i
    ' We have to be careful with this check because it is called on
    ' load when playing campaign on droids that are on the other
    ' map during missions, and those maps are usually larger.
    If Not worldOnMap(orig.x, orig.y) Or Not worldOnMap(dest.x, dest.y) Then
        fpathCheck = False
        Exit Function
    End If

    p = findNonblockingPosition(orig, propulsion)
    origTile = worldTile(p.x, p.y)
    p = findNonblockingPosition(dest, propulsion)
    destTile = worldTile(p.x, p.y)

    If Not (propulsion <> PROPULSION_TYPE_NUM) Then ASSERT_FAILURE "propulsion != PROPULSION_TYPE_NUM", "src/fpath.cpp:765", "fpathCheck", "Bad propulsion type": fpathCheck = False: Exit Function
    If Not (origTile <> -1 And destTile <> -1) Then ASSERT_FAILURE "origTile != nullptr && destTile != nullptr", "src/fpath.cpp:766", "fpathCheck", "Bad tile parameter": fpathCheck = False: Exit Function

    Select Case propulsion
    Case PROPULSION_TYPE_PROPELLOR, PROPULSION_TYPE_WHEELED, PROPULSION_TYPE_TRACKED, PROPULSION_TYPE_LEGGED, PROPULSION_TYPE_HALF_TRACKED
        fpathCheck = psMapTiles(origTile).limitedContinent = psMapTiles(destTile).limitedContinent
        Exit Function
    Case PROPULSION_TYPE_HOVER
        fpathCheck = psMapTiles(origTile).hoverContinent = psMapTiles(destTile).hoverContinent
        Exit Function
    Case PROPULSION_TYPE_LIFT
        fpathCheck = True                   ' assume no map uses skyscrapers to isolate areas
        Exit Function
    End Select

    ASSERT_FAILURE "false", "src/fpath.cpp:784", "fpathCheck", "Should never get here, unknown propulsion !"
    fpathCheck = False                      ' should never get here
End Function

' (VB, for the tests) fpathRoute, which fpathTest calls through fpathSimpleRoute
Public Function fpathRoute_test(psMove As MOVE_CONTROL, ByVal id As Long, ByVal startX As Long, ByVal startY As Long, ByVal tX As Long, ByVal tY As Long, _
                                ByVal propulsionType As Long, ByVal droidType As Long, ByVal moveType As Long, ByVal owner As Long, _
                                ByVal acceptNearest As Boolean, dstStructure As StructureBounds) As FPATH_RETVAL
    fpathRoute_test = fpathRoute(psMove, id, startX, startY, tX, tY, propulsionType, droidType, moveType, owner, acceptNearest, dstStructure)
End Function
