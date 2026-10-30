Attribute VB_Name = "move_c"
' src/move.cpp, move.h: Routines for moving units about the map.
' moveSteeringManager's static SteeringManager is steering_c's, given its
' CollisionAvoidanceBehavior on first use. Each function's static GridList
' is an array of this module. The callbacks are numbers: rayCast's
' (RAY_CALLBACK_MOVE_BLOCKING_TILE, its data this module's blockingData)
' and the AUDIO_CALLBACKs (CB_moveCheckDroidMovingAndVisible,
' CB_moveDroidStartCallback). moveOpenGates(DROID *, Vector2i) is
' moveOpenGates_tile. A Vector2i comparison is written out per coordinate.
Option Explicit

' lib/netplay/netplay.h: enum SYNC_OPT_TYPES
Private Const SYNC_OPT_FORMATION_SPEED_LIMITING As Long = 1

' max and min vtol heights above terrain
Private Const VTOL_HEIGHT_MIN As Long = 250
Private Const VTOL_HEIGHT_LEVEL As Long = 300
Private Const VTOL_HEIGHT_MAX As Long = 350

' Maximum size of an object for collision
Private Const OBJ_MAXRADIUS As Long = TILE_UNITS * 4

' how long a shuffle can propagate before they all stop
Private Const MOVE_SHUFFLETIME As Long = 10000

' Length of time a droid has to be stationery to be considered blocked
Private Const BLOCK_TIME As Long = 6000
Private Const SHUFFLE_BLOCK_TIME As Long = 2000
' How long a droid has to be stationary before stopping trying to move
Private Const BLOCK_PAUSETIME As Long = 1500
Private Const BLOCK_PAUSERELEASE As Long = 500
' How far a droid has to move before it is no longer 'stationary'
Private Const BLOCK_DIST As Long = 64
' How far a droid has to rotate before it is no longer 'stationary'
Private Const BLOCK_DIR As Long = 90

' How far out from an obstruction to start avoiding it
Private Const AVOID_DIST As Long = TILE_UNITS * 2

' Speed to approach a final way point, if possible.
Private Const MIN_END_SPEED As Long = 60

' distance from final way point to start slowing
Private Const END_SPEED_RANGE As Long = 3 * TILE_UNITS

' how long to pause after firing a FOM_NO weapon
Private Const FOM_MOVEPAUSE As Long = 1500

' distance to consider droids for a shuffle
Private Const SHUFFLE_DIST As Long = 3 * TILE_UNITS \ 2
' how far to move for a shuffle
Private Const SHUFFLE_MOVE As Long = 2 * TILE_UNITS \ 2

' Extra precision added to movement calculations.
Private Const EXTRA_BITS As Long = 8
Private Const EXTRA_PRECISION As Long = 256    ' (1 << EXTRA_BITS)

Private Const PITCH_LIMIT As Long = 150

Private Const MAX_SPEED_PITCH As Long = 60

' Watermelon:fix these magic number...the collision radius should be based on pie imd radius not some static int's...
Private Const mvPersRad As Long = 20
Private Const mvCybRad As Long = 30
Private Const mvSmRad As Long = 40
Private Const mvMedRad As Long = 50
Private Const mvLgRad As Long = 60

' checkLocalFeatures: scan the neighbours
Private Const DROIDDIST As Long = (TILE_UNITS * 5) \ 2
Private Const MAX_PICKUP_DISTANCE As Long = TILE_UNITS \ 2

' (VB) the AUDIO_CALLBACKs of this module (numbers, as CB_droidAudioTrackStopped)
Public Const CB_moveCheckDroidMovingAndVisible As Long = 2
Public Const CB_moveDroidStartCallback As Long = 3

Private Type BLOCKING_CALLBACK_DATA
    propulsionType As Long              ' PROPULSION_TYPE
    blocking As Boolean
    src As Vector2i
    dst As Vector2i
End Type

Private playerFormationSpeedLimiting() As Boolean     ' std::vector<bool>

Private blockingData As BLOCKING_CALLBACK_DATA         ' moveDirectPathToWaypoint's data, read by moveBlockingTileCallback

' the functions' static GridLists (static to avoid allocations)
Private gridListShuffle() As Long       ' moveShuffleDroid
Private gridListSquish() As Long        ' moveCheckSquished
Private gridListSlide() As Long         ' moveCalcDroidSlide
Private gridListLocal() As Long         ' checkLocalFeatures

' static steering::SteeringManager& moveSteeringManager(): the instance is made on first use
Private Sub moveSteeringManager()
    If Not SteeringManager_ready() Then
        SteeringManager_addBehavior STEERING_BEHAVIOR_COLLISION_AVOIDANCE
    End If
End Sub

Public Sub moveInit()
    Dim i As Long
    ' Initialize formation speed limiting to off for all players
    ReDim playerFormationSpeedLimiting(0 To MAX_PLAYERS - 1)
    For i = 0 To MAX_PLAYERS - 1
        playerFormationSpeedLimiting(i) = False
    Next
End Sub

Public Function moveSetFormationSpeedLimiting(ByVal player As Long, ByVal enabled As Boolean) As Boolean
    Dim currentPlayer As Long, optType As Long, value As Long, w As MessageWriter
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/move.cpp:131", "moveSetFormationSpeedLimiting", "Invalid player: %u", player: moveSetFormationSpeedLimiting = False: Exit Function
    If myResponsibility(player) Then
        ' always send a net message! (must be synchronized!)
        If Not (ULT(player, 255)) Then ASSERT_FAILURE "player < static_cast<uint32_t>(std::numeric_limits<uint8_t>::max())", "src/move.cpp:135", "moveSetFormationSpeedLimiting", "player exceeds expected bounds?": moveSetFormationSpeedLimiting = False: Exit Function
        currentPlayer = player And &HFF&
        optType = SYNC_OPT_FORMATION_SPEED_LIMITING
        If enabled Then value = 1 Else value = 0
        Set w = NETbeginEncode(NETgameQueue(currentPlayer), GAME_SYNC_OPT_CHANGE)
        NETuint8_t w, currentPlayer         ' player
        NETuint8_t w, optType
        NETuint8_t w, value                 ' formation speed limiting value
        moveSetFormationSpeedLimiting = NETend(w)
    Else
        moveSetFormationSpeedLimiting = False
    End If
End Function

' bool moveToggleFormationSpeedLimiting(uint32_t player, bool *pBoolResultingValue): hasResultingValue says whether
' C++ had a pointer
Public Function moveToggleFormationSpeedLimiting(ByVal player As Long, Optional ByVal hasResultingValue As Boolean = False, Optional pBoolResultingValue As Boolean) As Boolean
    Dim newValue As Boolean
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/move.cpp:153", "moveToggleFormationSpeedLimiting", "Invalid player: %u", player: moveToggleFormationSpeedLimiting = False: Exit Function
    If Not (myResponsibility(player)) Then ASSERT_FAILURE "myResponsibility(player)", "src/move.cpp:154", "moveToggleFormationSpeedLimiting", "Cannot change for player: %u", player: moveToggleFormationSpeedLimiting = False: Exit Function
    newValue = Not moveFormationSpeedLimitingOn(player)
    If moveSetFormationSpeedLimiting(player, newValue) Then
        If hasResultingValue Then
            pBoolResultingValue = newValue
        End If
        moveToggleFormationSpeedLimiting = True
        Exit Function
    End If
    moveToggleFormationSpeedLimiting = False
End Function

Public Function moveFormationSpeedLimitingOn(ByVal player As Long) As Boolean
    moveFormationSpeedLimitingOn = playerFormationSpeedLimiting(player)
End Function

Public Function recvSyncOptChange(ByVal queue As Long) As Boolean
    Dim player As Long, optType As Long, value As Long, r As MessageReader

    Set r = NETbeginDecode(queue, GAME_SYNC_OPT_CHANGE)
    NETuint8_t r, player    ' the player
    NETuint8_t r, optType
    NETuint8_t r, value
    NETend r

    If Not canGiveOrdersFor(NETQUEUE_index(queue), player) Then
        debug_ 186, LOG_WARNING, "recvSyncOptChange", "Sync opt change (by %d) for wrong player (%d).", NETQUEUE_index(queue), player
        syncDebug "recvSyncOptChange", "Wrong player."
        recvSyncOptChange = False
        Exit Function
    End If

    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/move.cpp:191", "recvSyncOptChange", "Invalid player: %u", player: recvSyncOptChange = False: Exit Function

    Select Case optType
    Case SYNC_OPT_FORMATION_SPEED_LIMITING
        debug_ 196, LOG_WZ, "recvSyncOptChange", "Received formation speed limiting change for player: %d, from: %d", player, NETQUEUE_index(queue)
        playerFormationSpeedLimiting(player) = (value <> 0)
    Case Else
        debug_ 200, LOG_WARNING, "recvSyncOptChange", "Sync opt change for player: %d, invalid sync opt type: %d", player, optType
        syncDebug "recvSyncOptChange", "Invalid sync opt type: %d", optType
        recvSyncOptChange = False
        Exit Function
    End Select

    recvSyncOptChange = True
End Function

Public Function moveDescription(ByVal status As MOVE_STATUS) As String
    Select Case status
    Case MOVEINACTIVE: moveDescription = "Inactive"
    Case MOVENAVIGATE: moveDescription = "Navigate"
    Case MOVETURN: moveDescription = "Turn"
    Case MOVEPAUSE: moveDescription = "Pause"
    Case MOVEPOINTTOPOINT: moveDescription = "P2P"
    Case MOVETURNTOTARGET: moveDescription = "Turn2target"
    Case MOVEHOVER: moveDescription = "Hover"
    Case MOVEWAITROUTE: moveDescription = "Waitroute"
    Case MOVESHUFFLE: moveDescription = "Shuffle"
    Case Else: moveDescription = "Error"    ' satisfy compiler
    End Select
End Function

' Set a target location in world coordinates for a droid to move to
' @return true if the routing was successful, if false then the calling code
'         should not try to route here again for a while
Private Function moveDroidToBase(ByVal psDroid As Long, ByVal x As Long, ByVal y As Long, ByVal bFormation As Boolean, ByVal moveType As Long) As Boolean
    Dim retVal As FPATH_RETVAL, psFormation As Long, f As Long, n As Long
    Dim fmx1 As Long, fmy1 As Long, fmx2 As Long, fmy2 As Long

    retVal = FPR_OK

    checkDroid psDroid, "src/move.cpp:238", "moveDroidToBase", max_check_object_recursion

    ' in multiPlayer make Transporter move like the vtols
    If DROID_isTransporter(psDroid) And game.maxPlayers = 0 Then
        fpathSetDirectRoute psDroid, x, y
        aDroid(psDroid).sMove.Status = MOVENAVIGATE
        aDroid(psDroid).sMove.pathIndex = 0
        moveDroidToBase = True
        Exit Function
    ' NOTE: While Vtols can fly, then can't go through things, like the transporter.
    ElseIf (game.maxPlayers > 0 And DROID_isTransporter(psDroid)) Then
        fpathSetDirectRoute psDroid, x, y
        retVal = FPR_OK
    Else
        retVal = fpathDroidRoute(psDroid, x, y, moveType)
    End If

    If retVal = FPR_OK Then
        ' bit of a hack this - john
        ' if astar doesn't have a complete route, it returns a route to the nearest clear tile.
        ' the location of the clear tile is in DestinationX,DestinationY.
        ' reset x,y to this position so the formation gets set up correctly
        x = aDroid(psDroid).sMove.destination.x
        y = aDroid(psDroid).sMove.destination.y

        objTrace aDroid(psDroid).id, "moveDroidToBase", "unit %d: path ok - base Speed %u, speed %d, target(%u|%d, %u|%d)", _
                 aDroid(psDroid).id, aDroid(psDroid).baseSpeed, aDroid(psDroid).sMove.speed, x, map_coord(x), y, map_coord(y)

        aDroid(psDroid).sMove.Status = MOVENAVIGATE
        aDroid(psDroid).sMove.pathIndex = 0

        ' leave any old formation
        If aDroid(psDroid).sMove.psFormation <> 0 Then
            formationLeave aDroid(psDroid).sMove.psFormation, psDroid
            aDroid(psDroid).sMove.psFormation = 0
        End If

        If bFormation Then
            ' join a formation if it exists at the destination
            psFormation = formationFind(aDroid(psDroid).player, x, y)

            If psFormation <> 0 Then
                aDroid(psDroid).sMove.psFormation = psFormation
                formationJoin psFormation, psDroid
            Else
                n = aDroid(psDroid).sMove.asPath_size
                If Not (n > 0) Then ASSERT_FAILURE "psDroid->sMove.asPath.size() > 0", "src/move.cpp:294", "moveDroidToBase", "Invalid point count for path of %s", droidGetName(psDroid): moveDroidToBase = False: Exit Function
                ' align the formation with the last path of the route
                fmx2 = aDroid(psDroid).sMove.asPath(n - 1).x
                fmy2 = aDroid(psDroid).sMove.asPath(n - 1).y
                If n = 1 Then
                    fmx1 = aDroid(psDroid).pos.x
                    fmy1 = aDroid(psDroid).pos.y
                Else
                    If Not (n > 1) Then ASSERT_FAILURE "psDroid->sMove.asPath.size() > 1", "src/move.cpp:305", "moveDroidToBase", "Invalid point count for path of %s", droidGetName(psDroid): moveDroidToBase = False: Exit Function
                    fmx1 = aDroid(psDroid).sMove.asPath(n - 2).x
                    fmy1 = aDroid(psDroid).sMove.asPath(n - 2).y
                End If

                ' no formation so create a new one
                f = aDroid(psDroid).sMove.psFormation
                If formationNew(f, aDroid(psDroid).player, FT_LINE, x, y, calcDirection(fmx1, fmy1, fmx2, fmy2)) Then
                    aDroid(psDroid).sMove.psFormation = f
                    formationJoin aDroid(psDroid).sMove.psFormation, psDroid
                End If
            End If
        End If
    ElseIf retVal = FPR_WAIT Then
        ' the route will be calculated by the path-finding thread
        aDroid(psDroid).sMove.Status = MOVEWAITROUTE
        aDroid(psDroid).sMove.destination.x = x
        aDroid(psDroid).sMove.destination.y = y
    Else ' if (retVal == FPR_FAILED)
        objTrace aDroid(psDroid).id, "moveDroidToBase", "Path to (%d, %d) failed for droid %d", x, y, aDroid(psDroid).id
        aDroid(psDroid).sMove.Status = MOVEINACTIVE
        actionDroid psDroid, DACTION_SULK
        moveDroidToBase = False
        Exit Function
    End If

    checkDroid psDroid, "src/move.cpp:334", "moveDroidToBase", max_check_object_recursion
    moveDroidToBase = True
End Function

' Move a droid to a location, joining a formation
' @see moveDroidToBase() for the parameter and return value specification
Public Function moveDroidTo(ByVal psDroid As Long, ByVal x As Long, ByVal y As Long, Optional ByVal moveType As Long = FMT_MOVE) As Boolean
    moveDroidTo = moveDroidToBase(psDroid, x, y, True, moveType)
End Function

' Move a droid to a location, not joining a formation
' @see moveDroidToBase() for the parameter and return value specification
Public Function moveDroidToNoFormation(ByVal psDroid As Long, ByVal x As Long, ByVal y As Long, Optional ByVal moveType As Long = FMT_MOVE) As Boolean
    If Not (x <> 0 And y <> 0) Then ASSERT_FAILURE "x > 0 && y > 0", "src/move.cpp:351", "moveDroidToNoFormation", "Bad movement position": moveDroidToNoFormation = False: Exit Function
    moveDroidToNoFormation = moveDroidToBase(psDroid, x, y, False, moveType)
End Function

' Move a droid directly to a location.
' @note This is (or should be) used for VTOLs only.
Public Sub moveDroidToDirect(ByVal psDroid As Long, ByVal x As Long, ByVal y As Long)
    Dim ok As Boolean
    ok = psDroid <> 0
    If ok Then ok = DROID_isVtol(psDroid)
    If Not ok Then ASSERT_FAILURE "psDroid != nullptr && psDroid->isVtol()", "src/move.cpp:361", "moveDroidToDirect", "Only valid for a VTOL unit": Exit Sub

    fpathSetDirectRoute psDroid, x, y
    aDroid(psDroid).sMove.Status = MOVENAVIGATE
    aDroid(psDroid).sMove.pathIndex = 0

    ' leave any old formation
    If aDroid(psDroid).sMove.psFormation <> 0 Then
        formationLeave aDroid(psDroid).sMove.psFormation, psDroid
        aDroid(psDroid).sMove.psFormation = 0
    End If
End Sub

' Turn a droid towards a given location.
Public Sub moveTurnDroid(ByVal psDroid As Long, ByVal x As Long, ByVal y As Long)
    Dim moveDir As Long
    moveDir = calcDirection(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, x, y)

    If aDroid(psDroid).rot.direction <> moveDir Then
        aDroid(psDroid).sMove.target.x = x
        aDroid(psDroid).sMove.target.y = y
        aDroid(psDroid).sMove.Status = MOVETURNTOTARGET
    End If
End Sub

' Tell a droid to move out the way for a shuffle
Private Sub moveShuffleDroid(ByVal psDroid As Long, s As Vector2i)
    Dim mx As Long, my As Long
    Dim frontClear As Boolean, leftClear As Boolean, rightClear As Boolean
    Dim lvx As Long, lvy As Long, rvx As Long, rvy As Long, svx As Long, svy As Long
    Dim shuffleMove As Long, shuffleDir As Long, shuffleMag As Long, droidPropType As Long
    Dim n As Long, gi As Long, psCurr As Long, droidDir As Long, diff As Long, tar As Vector2i

    frontClear = True
    leftClear = True
    rightClear = True

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/move.cpp:398", "moveShuffleDroid", "Bad droid pointer": Exit Sub
    checkDroid psDroid, "src/move.cpp:399", "moveShuffleDroid", max_check_object_recursion

    shuffleDir = iAtan2(s.x, s.y)
    shuffleMag = iHypot(s.x, s.y)

    If shuffleMag = 0 Then
        Exit Sub
    End If

    shuffleMove = SHUFFLE_MOVE

    ' calculate the possible movement vectors
    svx = s.x * shuffleMove \ shuffleMag  ' Straight in the direction of s.
    svy = s.y * shuffleMove \ shuffleMag

    lvx = -svy  ' 90 degrees to the... right?
    lvy = svx

    rvx = svy   ' 90 degrees to the... left?
    rvy = -svx

    droidPropType = asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType
    ' check for blocking tiles
    If fpathBlockingTile(map_coord(aDroid(psDroid).pos.x + lvx), _
                         map_coord(aDroid(psDroid).pos.y + lvy), droidPropType) Then
        leftClear = False
    ElseIf fpathBlockingTile(map_coord(aDroid(psDroid).pos.x + rvx), _
                             map_coord(aDroid(psDroid).pos.y + rvy), droidPropType) Then
        rightClear = False
    ElseIf fpathBlockingTile(map_coord(aDroid(psDroid).pos.x + svx), _
                             map_coord(aDroid(psDroid).pos.y + svy), droidPropType) Then
        frontClear = False
    End If

    ' find any droids that could block the shuffle
    n = gridStartIterate(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, SHUFFLE_DIST, gridListShuffle)
    For gi = 0 To n - 1
        psCurr = castDroid(gridListShuffle(gi))
        If psCurr = 0 Then
            GoTo continue_
        End If
        If aDroid(psCurr).died <> 0 Or psCurr = psDroid Then
            GoTo continue_
        End If

        droidDir = iAtan2(aDroid(psCurr).pos.x - aDroid(psDroid).pos.x, aDroid(psCurr).pos.y - aDroid(psDroid).pos.y)
        diff = angleDelta(shuffleDir - droidDir)
        If diff > -DEG(135) And diff < -DEG(45) Then
            leftClear = False
        ElseIf diff > DEG(45) And diff < DEG(135) Then
            rightClear = False
        End If
continue_:
    Next

    ' calculate a target
    If leftClear Then
        mx = lvx
        my = lvy
    ElseIf rightClear Then
        mx = rvx
        my = rvy
    ElseIf frontClear Then
        mx = svx
        my = svy
    Else
        ' nowhere to shuffle to, quit
        Exit Sub
    End If

    ' check the location for vtols
    tar.x = aDroid(psDroid).pos.x + mx
    tar.y = aDroid(psDroid).pos.y + my
    If DROID_isVtol(psDroid) Then
        actionVTOLLandingPos psDroid, tar
    End If

    ' set up the move state
    If aDroid(psDroid).sMove.Status <> MOVESHUFFLE Then
        aDroid(psDroid).sMove.shuffleStart = gameTime
    End If
    aDroid(psDroid).sMove.Status = MOVESHUFFLE
    aDroid(psDroid).sMove.src.x = aDroid(psDroid).pos.x
    aDroid(psDroid).sMove.src.y = aDroid(psDroid).pos.y
    aDroid(psDroid).sMove.target = tar
    aDroid(psDroid).sMove.asPath_size = 0   ' asPath.clear()
    aDroid(psDroid).sMove.pathIndex = 0

    If aDroid(psDroid).sMove.psFormation <> 0 Then
        formationLeave aDroid(psDroid).sMove.psFormation, psDroid
        aDroid(psDroid).sMove.psFormation = 0
    End If

    checkDroid psDroid, "src/move.cpp:509", "moveShuffleDroid", max_check_object_recursion
End Sub

' Stop a droid from moving.
Public Sub moveStopDroid(ByVal psDroid As Long)
    Dim psPropStats As Long
    checkDroid psDroid, "src/move.cpp:517", "moveStopDroid", max_check_object_recursion
    psPropStats = DROID_getPropulsionStats(psDroid)
    If Not (psPropStats <> -1) Then ASSERT_FAILURE "psPropStats != nullptr", "src/move.cpp:519", "moveStopDroid", "invalid propulsion stats pointer": Exit Sub

    If asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_LIFT Then
        aDroid(psDroid).sMove.Status = MOVEHOVER
    Else
        aDroid(psDroid).sMove.Status = MOVEINACTIVE
    End If
End Sub

' Stops a droid dead in its tracks.
' Doesn't allow for any little skidding bits.
' @param psDroid the droid to stop from moving
Public Sub moveReallyStopDroid(ByVal psDroid As Long)
    checkDroid psDroid, "src/move.cpp:537", "moveReallyStopDroid", max_check_object_recursion

    aDroid(psDroid).sMove.Status = MOVEINACTIVE
    aDroid(psDroid).sMove.speed = 0
End Sub

' Get pitch and roll from direction and tile data
Public Sub updateDroidOrientation(ByVal psDroid As Long)
    Dim hx0 As Long, hx1 As Long, hy0 As Long, hy1 As Long
    Dim newPitch As Long, deltaPitch As Long, pitchLimit As Long
    Dim dzdx As Long, dzdy As Long, dzdv As Long, dzdw As Long
    Const d As Long = 20
    Dim vX As Long, vY As Long

    If aDroid(psDroid).droidType = DROID_PERSON Or DROID_isCyborg(psDroid) Or DROID_isFlightBasedTransporter(psDroid) _
        Or DROID_isFlying(psDroid) Then
        ' The ground doesn't affect the pitch/roll of these droids
        Exit Sub
    End If

    ' Find the height of 4 points around the droid.
    '    hy0
    ' hx0 * hx1      (* = droid)
    '    hy1
    hx1 = map_Height(aDroid(psDroid).pos.x + d, aDroid(psDroid).pos.y)
    hx0 = map_Height(MAX_(0, aDroid(psDroid).pos.x - d), aDroid(psDroid).pos.y)
    hy1 = map_Height(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y + d)
    hy0 = map_Height(aDroid(psDroid).pos.x, MAX_(0, aDroid(psDroid).pos.y - d))

    'update height in case were in the bottom of a trough
    aDroid(psDroid).pos.z = MAX_(aDroid(psDroid).pos.z, (hx0 + hx1) \ 2)
    aDroid(psDroid).pos.z = MAX_(aDroid(psDroid).pos.z, (hy0 + hy1) \ 2)

    ' Vector of length 65536 pointing in direction droid is facing.
    vX = iSin(aDroid(psDroid).rot.direction)
    vY = iCos(aDroid(psDroid).rot.direction)

    ' Calculate pitch of ground.
    dzdx = hx1 - hx0                                ' 2*d*dz(x, y)/dx       of ground
    dzdy = hy1 - hy0                                ' 2*d*dz(x, y)/dy       of ground
    dzdv = dzdx * vX + dzdy * vY                    ' 2*d*dz(x, y)/dv << 16 of ground, where v is the direction the droid is facing.
    newPitch = iAtan2(dzdv, (2 * d) * 65536)        ' pitch = atan(dz(x, y)/dv)/2pi << 16

    deltaPitch = angleDelta(newPitch - aDroid(psDroid).rot.pitch)

    ' Limit the rate the front comes down to simulate momentum
    pitchLimit = gameTimeAdjustedIncrement(DEG(PITCH_LIMIT))
    deltaPitch = MAX_(deltaPitch, -pitchLimit)

    ' Update pitch.
    aDroid(psDroid).rot.pitch = (aDroid(psDroid).rot.pitch + deltaPitch) And &HFFFF&

    ' Calculate and update roll of ground (not taking pitch into account, but good enough).
    dzdw = dzdx * vY - dzdy * vX                    ' 2*d*dz(x, y)/dw << 16 of ground, where w is at right angles to the direction the droid is facing.
    aDroid(psDroid).rot.roll = iAtan2(dzdw, (2 * d) * 65536) And &HFFFF&   ' pitch = atan(dz(x, y)/dw)/2pi << 16
End Sub

' rayCast's callback (RAY_CALLBACK_MOVE_BLOCKING_TILE); its data is blockingData
Public Function moveBlockingTileCallback(pos As Vector2i) As Boolean
    If Not blockingData.blocking Then
        If (pos.x <> blockingData.src.x Or pos.y <> blockingData.src.y) And (pos.x <> blockingData.dst.x Or pos.y <> blockingData.dst.y) Then
            blockingData.blocking = fpathBlockingTile(map_coord(pos.x), map_coord(pos.y), blockingData.propulsionType)
        End If
    End If
    moveBlockingTileCallback = Not blockingData.blocking
End Function

' Returns -1 - distance if the direct path to the waypoint is blocked, otherwise returns the distance to the waypoint.
Private Function moveDirectPathToWaypoint(ByVal psDroid As Long, ByVal positionIndex As Long) As Long
    Dim src As Vector2i, dst As Vector2i, delta As Vector2i, dist As Long
    src.x = aDroid(psDroid).pos.x
    src.y = aDroid(psDroid).pos.y
    dst = aDroid(psDroid).sMove.asPath(positionIndex)
    delta.x = dst.x - src.x
    delta.y = dst.y - src.y
    dist = iHypot(delta.x, delta.y)
    blockingData.propulsionType = asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType
    blockingData.blocking = False
    blockingData.src = src
    blockingData.dst = dst
    rayCast src, dst, RAY_CALLBACK_MOVE_BLOCKING_TILE
    If blockingData.blocking Then moveDirectPathToWaypoint = -1 - dist Else moveDirectPathToWaypoint = dist
End Function

' Returns true if still able to find the path.
Private Function moveBestTarget(ByVal psDroid As Long) As Boolean
    Dim positionIndex As Long, dist As Long
    positionIndex = MAX_(aDroid(psDroid).sMove.pathIndex - 1, 0)
    dist = moveDirectPathToWaypoint(psDroid, positionIndex)
    If dist >= 0 Then
        ' Look ahead in the path.
        Do While dist >= 0 And dist < TILE_UNITS * 5
            positionIndex = positionIndex + 1
            If positionIndex >= aDroid(psDroid).sMove.asPath_size Then
                dist = -1
                Exit Do  ' Reached end of path.
            End If
            dist = moveDirectPathToWaypoint(psDroid, positionIndex)
        Loop
        If dist < 0 Then
            positionIndex = positionIndex - 1
        End If
    Else
        ' Lost sight of path, backtrack.
        Do While dist < 0 And dist >= -TILE_UNITS * 7 And positionIndex > 0
            positionIndex = positionIndex - 1
            dist = moveDirectPathToWaypoint(psDroid, positionIndex)
        Loop
        If dist < 0 Then
            moveBestTarget = False  ' Couldn't find path, and backtracking didn't help.
            Exit Function
        End If
    End If
    aDroid(psDroid).sMove.pathIndex = positionIndex + 1
    aDroid(psDroid).sMove.src.x = aDroid(psDroid).pos.x
    aDroid(psDroid).sMove.src.y = aDroid(psDroid).pos.y
    aDroid(psDroid).sMove.target = aDroid(psDroid).sMove.asPath(positionIndex)
    moveBestTarget = True
End Function

' Get the next target point from the route
Private Function moveNextTarget(ByVal psDroid As Long) As Boolean
    checkDroid psDroid, "src/move.cpp:676", "moveNextTarget", max_check_object_recursion

    ' See if there is anything left in the move list
    If aDroid(psDroid).sMove.pathIndex = aDroid(psDroid).sMove.asPath_size Then
        moveNextTarget = False
        Exit Function
    End If
    If Not (aDroid(psDroid).sMove.pathIndex >= 0 And aDroid(psDroid).sMove.pathIndex < aDroid(psDroid).sMove.asPath_size) Then ASSERT_FAILURE "psDroid->sMove.pathIndex >= 0 && psDroid->sMove.pathIndex < (int)psDroid->sMove.asPath.size()", "src/move.cpp:683", "moveNextTarget", "psDroid->sMove.pathIndex out of bounds %d/%d.", aDroid(psDroid).sMove.pathIndex, aDroid(psDroid).sMove.asPath_size: moveNextTarget = False: Exit Function

    If aDroid(psDroid).sMove.pathIndex = 0 Then
        aDroid(psDroid).sMove.src.x = aDroid(psDroid).pos.x
        aDroid(psDroid).sMove.src.y = aDroid(psDroid).pos.y
    Else
        aDroid(psDroid).sMove.src = aDroid(psDroid).sMove.asPath(aDroid(psDroid).sMove.pathIndex - 1)
    End If
    aDroid(psDroid).sMove.target = aDroid(psDroid).sMove.asPath(aDroid(psDroid).sMove.pathIndex)
    aDroid(psDroid).sMove.pathIndex = aDroid(psDroid).sMove.pathIndex + 1

    checkDroid psDroid, "src/move.cpp:696", "moveNextTarget", max_check_object_recursion
    moveNextTarget = True
End Function

' Get the radius of a base object for collision
Public Function moveObjRadius(ByVal psObj As Long) As Long
    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        If aDroid(psObj).droidType = DROID_PERSON Then
            moveObjRadius = mvPersRad
        ElseIf DROID_isCyborg(psObj) Then
            moveObjRadius = mvCybRad
        Else
            Select Case asBodyStats(DROID_getBodyStats(psObj)).size
            Case SIZE_LIGHT
                moveObjRadius = mvSmRad

            Case SIZE_MEDIUM
                moveObjRadius = mvMedRad

            Case SIZE_HEAVY
                moveObjRadius = mvLgRad

            Case SIZE_SUPER_HEAVY
                moveObjRadius = 130

            Case Else
                moveObjRadius = iIMDBaseShape_radius(aDroid(psObj).sDisplay.imd)
            End Select
        End If

    Case OBJ_STRUCTURE
        moveObjRadius = iIMDBaseShape_radius(aStruct(psObj).sDisplay.imd) \ 2

    Case OBJ_FEATURE
        moveObjRadius = iIMDBaseShape_radius(aFeature(psObj).sDisplay.imd) \ 2

    Case Else
        ASSERT_FAILURE "false", "src/move.cpp:749", "moveObjRadius", "unknown object type"
        moveObjRadius = 0
    End Select
End Function

' see if a Droid has run over a person
Private Sub moveCheckSquished(ByVal psDroid As Long, ByVal emx As Long, ByVal emy As Long)
    Dim rad As Long, radSq As Long, objR As Long, xdiff As Long, ydiff As Long, distSq As Long
    Dim droidR As Long, mx As Long, my As Long, n As Long, gi As Long, psObj As Long

    droidR = moveObjRadius(psDroid)
    mx = gameTimeAdjustedAverage(emx, EXTRA_PRECISION)
    my = gameTimeAdjustedAverage(emy, EXTRA_PRECISION)

    n = gridStartIterate(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, OBJ_MAXRADIUS, gridListSquish)
    For gi = 0 To n - 1
        psObj = gridListSquish(gi)
        If SIMPLE_OBJECT_type(psObj) <> OBJ_DROID Then
            ' ignore everything but people
            GoTo continue_
        End If
        If aDroid(psObj).droidType <> DROID_PERSON Then
            GoTo continue_
        End If

        If Not (SIMPLE_OBJECT_type(psObj) = OBJ_DROID And aDroid(psObj).droidType = DROID_PERSON) Then ASSERT_FAILURE "psObj->type == OBJ_DROID && ((DROID *)psObj)->droidType == DROID_PERSON", "src/move.cpp:774", "moveCheckSquished", "squished - eerk": Exit Sub

        objR = moveObjRadius(psObj)
        rad = droidR + objR
        radSq = rad * rad

        xdiff = aDroid(psDroid).pos.x + mx - aDroid(psObj).pos.x
        ydiff = aDroid(psDroid).pos.y + my - aDroid(psObj).pos.y
        distSq = xdiff * xdiff + ydiff * ydiff

        If ((2 * radSq) \ 3) > distSq Then
            If (aDroid(psDroid).player <> aDroid(psObj).player) And Not aiCheckAlliances(aDroid(psDroid).player, aDroid(psObj).player) Then
                ' run over a bloke - kill him
                destroyDroid psObj, gameTime
                scoreUpdateVar WD_BARBARIANS_MOWED_DOWN
                giveExperienceForSquish psDroid
            End If
        End If
continue_:
    Next
End Sub

' See if the droid has been stopped long enough to give up on the move
Public Function moveBlocked(ByVal psDroid As Long) As Boolean
    Dim xdiff As Long, ydiff As Long, diffSq As Long, blockTime As Long

    If aDroid(psDroid).sMove.bumpTime = 0 Or ULT(gameTime, aDroid(psDroid).sMove.bumpTime) Then
        ' no bump - can't be blocked
        moveBlocked = False
        Exit Function
    End If

    ' See if the block can be cancelled
    If Abs(angleDelta(aDroid(psDroid).rot.direction - aDroid(psDroid).sMove.bumpDir)) > DEG(BLOCK_DIR) Then
        ' Move on, clear the bump
        aDroid(psDroid).sMove.bumpTime = 0
        aDroid(psDroid).sMove.lastBump = 0
        moveBlocked = False
        Exit Function
    End If
    xdiff = aDroid(psDroid).pos.x - aDroid(psDroid).sMove.bumpPos.x
    ydiff = aDroid(psDroid).pos.y - aDroid(psDroid).sMove.bumpPos.y
    diffSq = xdiff * xdiff + ydiff * ydiff
    If diffSq > BLOCK_DIST * BLOCK_DIST Then
        ' Move on, clear the bump
        aDroid(psDroid).sMove.bumpTime = 0
        aDroid(psDroid).sMove.lastBump = 0
        moveBlocked = False
        Exit Function
    End If

    If aDroid(psDroid).sMove.Status = MOVESHUFFLE Then
        blockTime = SHUFFLE_BLOCK_TIME
    Else
        blockTime = BLOCK_TIME
    End If

    If ULT(blockTime, gameTime - aDroid(psDroid).sMove.bumpTime) Then
        ' Stopped long enough - blocked
        aDroid(psDroid).sMove.bumpTime = 0
        aDroid(psDroid).sMove.lastBump = 0
        If Not isHumanPlayer(aDroid(psDroid).player) And bMultiPlayer Then
            aDroid(psDroid).lastFrustratedTime = gameTime
            objTrace aDroid(psDroid).id, "moveBlocked", "FRUSTRATED"
        Else
            objTrace aDroid(psDroid).id, "moveBlocked", "BLOCKED"
        End If
        ' if the unit cannot see the next way point - reroute it's got stuck
        If (bMultiPlayer Or aDroid(psDroid).player = selectedPlayer Or aDroid(psDroid).lastFrustratedTime = gameTime) _
           And aDroid(psDroid).sMove.pathIndex <> aDroid(psDroid).sMove.asPath_size Then
            objTrace aDroid(psDroid).id, "moveBlocked", "Trying to reroute to (%d,%d)", aDroid(psDroid).sMove.destination.x, aDroid(psDroid).sMove.destination.y
            moveDroidTo psDroid, aDroid(psDroid).sMove.destination.x, aDroid(psDroid).sMove.destination.y
            moveBlocked = False
            Exit Function
        End If

        moveBlocked = True
        Exit Function
    End If

    moveBlocked = False
End Function

' Calculate the actual movement to slide around
Private Sub moveCalcSlideVector(ByVal psDroid As Long, ByVal objX As Long, ByVal objY As Long, pMx As Long, pMy As Long)
    Dim dirX As Long, dirY As Long, dirMagSq As Long, dotRes As Long
    Dim mx As Long, my As Long, obstX As Long, obstY As Long
    mx = pMx
    my = pMy
    ' Calculate the vector to the obstruction
    obstX = aDroid(psDroid).pos.x - objX
    obstY = aDroid(psDroid).pos.y - objY

    ' if the target dir is the same, don't need to slide
    If obstX * mx + obstY * my >= 0 Then
        Exit Sub
    End If

    ' Choose the tangent vector to this on the same side as the target
    dotRes = obstY * mx - obstX * my
    If dotRes >= 0 Then
        dirX = obstY
        dirY = -obstX
    Else
        dirX = -obstY
        dirY = obstX
        dotRes = -dotRes
    End If
    dirMagSq = MAX_(1, dirX * dirX + dirY * dirY)

    ' Calculate the component of the movement in the direction of the tangent vector
    pMx = Fix(CDbl(dirX) * dotRes / dirMagSq)
    pMy = Fix(CDbl(dirY) * dotRes / dirMagSq)
End Sub

' moveOpenGates(DROID *psDroid, Vector2i tile)
Private Sub moveOpenGates_tile(ByVal psDroid As Long, tile As Vector2i)
    Dim psTile As Long, psObject As Long
    ' is the new tile a gate?
    If Not worldOnMap(tile.x, tile.y) Then
        Exit Sub
    End If
    psTile = mapTile(tile.x, tile.y)
    If Not DROID_isFlying(psDroid) And Not DROID_isFlightBasedTransporter(psDroid) Then
        psObject = psMapTiles(psTile).psObject
        If psObject <> 0 Then
            If SIMPLE_OBJECT_type(psObject) = OBJ_STRUCTURE Then
                If aiCheckAlliances(aStruct(psObject).player, aDroid(psDroid).player) Then
                    requestOpenGate psObject  ' If it's a friendly gate, open it. (It would be impolite to open an enemy gate.)
                End If
            End If
        End If
    End If
End Sub

Private Sub moveOpenGates(ByVal psDroid As Long)
    Dim r As Long, pos As Vector2i, tile As Vector2i
    r = aDroid(psDroid).sMove.speed * SAS_OPEN_SPEED \ GAME_TICKS_PER_SEC
    pos.x = aDroid(psDroid).pos.x + iSinR(aDroid(psDroid).sMove.moveDir, r)
    pos.y = aDroid(psDroid).pos.y + iCosR(aDroid(psDroid).sMove.moveDir, r)
    tile.x = map_coord(pos.x)
    tile.y = map_coord(pos.y)
    moveOpenGates_tile psDroid, tile
End Sub

' see if a droid has run into a blocking tile
' TODO See if this function can be simplified.
Private Sub moveCalcBlockingSlide(ByVal psDroid As Long, pmx As Long, pmy As Long, ByVal tarDir As Long, pSlideDir As Long)
    Dim propulsion As Long, horizX As Long, horizY As Long, vertX As Long, vertY As Long, slideDir As Long
    Dim mx As Long, my As Long, tx As Long, ty As Long, nx As Long, ny As Long, ntx As Long, nty As Long
    Dim blkCX As Long, blkCY As Long, ntile As Vector2i
    Dim intx As Long, inty As Long, bJumped As Boolean, jumpx As Long, jumpy As Long

    propulsion = asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType
    ' calculate the new coords and see if they are on a different tile
    mx = gameTimeAdjustedAverage(pmx, EXTRA_PRECISION)
    my = gameTimeAdjustedAverage(pmy, EXTRA_PRECISION)
    tx = map_coord(aDroid(psDroid).pos.x)
    ty = map_coord(aDroid(psDroid).pos.y)
    nx = aDroid(psDroid).pos.x + mx
    ny = aDroid(psDroid).pos.y + my
    ntx = map_coord(nx)
    nty = map_coord(ny)
    blkCX = world_coord(ntx) + TILE_UNITS \ 2
    blkCY = world_coord(nty) + TILE_UNITS \ 2

    checkDroid psDroid, "src/move.cpp:944", "moveCalcBlockingSlide", max_check_object_recursion

    ' is the new tile a gate?
    ntile.x = ntx
    ntile.y = nty
    moveOpenGates_tile psDroid, ntile

    ' is the new tile blocking?
    If Not fpathBlockingTile(ntx, nty, propulsion) Then
        ' not blocking, don't change the move vector
        Exit Sub
    End If

    ' if the droid is shuffling - just stop
    If aDroid(psDroid).sMove.Status = MOVESHUFFLE Then
        objTrace aDroid(psDroid).id, "moveCalcBlockingSlide", "Was shuffling, now stopped"
        aDroid(psDroid).sMove.Status = MOVEINACTIVE
    End If

    ' note the bump time and position if necessary
    If Not DROID_isVtol(psDroid) And _
       aDroid(psDroid).sMove.bumpTime = 0 Then
        aDroid(psDroid).sMove.bumpTime = gameTime
        aDroid(psDroid).sMove.lastBump = 0
        aDroid(psDroid).sMove.pauseTime = 0
        aDroid(psDroid).sMove.bumpPos = aDroid(psDroid).pos
        aDroid(psDroid).sMove.bumpDir = aDroid(psDroid).rot.direction
    End If

    If tx <> ntx And ty <> nty Then
        ' moved diagonally

        ' figure out where the other two possible blocking tiles are
        If mx < 0 Then horizX = ntx + 1 Else horizX = ntx - 1
        horizY = nty

        vertX = ntx
        If my < 0 Then vertY = nty + 1 Else vertY = nty - 1

        If fpathBlockingTile(horizX, horizY, propulsion) And fpathBlockingTile(vertX, vertY, propulsion) Then
            ' in a corner - choose an arbitrary slide
            If gameRand(2) = 0 Then
                pmx = 0
                pmy = -pmy
            Else
                pmx = -pmx
                pmy = 0
            End If
        ElseIf fpathBlockingTile(horizX, horizY, propulsion) Then
            pmy = 0
        ElseIf fpathBlockingTile(vertX, vertY, propulsion) Then
            pmx = 0
        Else
            moveCalcSlideVector psDroid, blkCX, blkCY, pmx, pmy
        End If
    ElseIf tx <> ntx Then
        ' moved horizontally - see which half of the tile were in
        If (aDroid(psDroid).pos.y And TILE_MASK) > TILE_UNITS \ 2 Then
            ' top half
            If fpathBlockingTile(ntx, nty + 1, propulsion) Then
                pmx = 0
            Else
                moveCalcSlideVector psDroid, blkCX, blkCY, pmx, pmy
            End If
        Else
            ' bottom half
            If fpathBlockingTile(ntx, nty - 1, propulsion) Then
                pmx = 0
            Else
                moveCalcSlideVector psDroid, blkCX, blkCY, pmx, pmy
            End If
        End If
    ElseIf ty <> nty Then
        ' moved vertically
        If (aDroid(psDroid).pos.x And TILE_MASK) > TILE_UNITS \ 2 Then
            ' top half
            If fpathBlockingTile(ntx + 1, nty, propulsion) Then
                pmy = 0
            Else
                moveCalcSlideVector psDroid, blkCX, blkCY, pmx, pmy
            End If
        Else
            ' bottom half
            If fpathBlockingTile(ntx - 1, nty, propulsion) Then
                pmy = 0
            Else
                moveCalcSlideVector psDroid, blkCX, blkCY, pmx, pmy
            End If
        End If
    Else ' if (tx == ntx && ty == nty)
        ' on a blocking tile - see if we need to jump off
        intx = aDroid(psDroid).pos.x And TILE_MASK
        inty = aDroid(psDroid).pos.y And TILE_MASK
        bJumped = False
        jumpx = aDroid(psDroid).pos.x
        jumpy = aDroid(psDroid).pos.y

        If intx < TILE_UNITS \ 2 Then
            If inty < TILE_UNITS \ 2 Then
                ' top left
                If (mx < 0) And fpathBlockingTile(tx - 1, ty, propulsion) Then
                    bJumped = True
                    jumpy = (jumpy And Not TILE_MASK) - 1
                End If
                If (my < 0) And fpathBlockingTile(tx, ty - 1, propulsion) Then
                    bJumped = True
                    jumpx = (jumpx And Not TILE_MASK) - 1
                End If
            Else
                ' bottom left
                If (mx < 0) And fpathBlockingTile(tx - 1, ty, propulsion) Then
                    bJumped = True
                    jumpy = (jumpy And Not TILE_MASK) + TILE_UNITS
                End If
                If (my >= 0) And fpathBlockingTile(tx, ty + 1, propulsion) Then
                    bJumped = True
                    jumpx = (jumpx And Not TILE_MASK) - 1
                End If
            End If
        Else
            If inty < TILE_UNITS \ 2 Then
                ' top right
                If (mx >= 0) And fpathBlockingTile(tx + 1, ty, propulsion) Then
                    bJumped = True
                    jumpy = (jumpy And Not TILE_MASK) - 1
                End If
                If (my < 0) And fpathBlockingTile(tx, ty - 1, propulsion) Then
                    bJumped = True
                    jumpx = (jumpx And Not TILE_MASK) + TILE_UNITS
                End If
            Else
                ' bottom right
                If (mx >= 0) And fpathBlockingTile(tx + 1, ty, propulsion) Then
                    bJumped = True
                    jumpy = (jumpy And Not TILE_MASK) + TILE_UNITS
                End If
                If (my >= 0) And fpathBlockingTile(tx, ty + 1, propulsion) Then
                    bJumped = True
                    jumpx = (jumpx And Not TILE_MASK) + TILE_UNITS
                End If
            End If
        End If

        If bJumped Then
            aDroid(psDroid).pos.x = MAX_(0, jumpx)
            aDroid(psDroid).pos.y = MAX_(0, jumpy)
            pmx = 0
            pmy = 0
        Else
            moveCalcSlideVector psDroid, blkCX, blkCY, pmx, pmy
        End If
    End If

    slideDir = iAtan2(pmx, pmy)
    If ntx <> tx Then
        ' hit a horizontal block
        If (tarDir < DEG(90) Or tarDir > DEG(270)) And _
           (slideDir >= DEG(90) And slideDir <= DEG(270)) Then
            slideDir = tarDir
        ElseIf (tarDir >= DEG(90) And tarDir <= DEG(270)) And _
               (slideDir < DEG(90) Or slideDir > DEG(270)) Then
            slideDir = tarDir
        End If
    End If
    If nty <> ty Then
        ' hit a vertical block
        If (tarDir < DEG(180)) And _
           (slideDir >= DEG(180)) Then
            slideDir = tarDir
        ElseIf (tarDir >= DEG(180)) And _
               (slideDir < DEG(180)) Then
            slideDir = tarDir
        End If
    End If
    pSlideDir = slideDir

    checkDroid psDroid, "src/move.cpp:1184", "moveCalcBlockingSlide", max_check_object_recursion
End Sub

' see if a droid has run into another droid
' Only consider stationery droids
Private Sub moveCalcDroidSlide(ByVal psDroid As Long, pmx As Long, pmy As Long)
    Dim droidR As Long, rad As Long, radSq As Long, objR As Long, xdiff As Long, ydiff As Long, distSq As Long, spmx As Long, spmy As Long
    Dim bLegs As Boolean, psObst As Long, n As Long, gi As Long, psObj As Long, psObjcast As Long, psShuffleDroid As Long
    Dim overFlying As Boolean, s As Vector2i

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/move.cpp:1194", "moveCalcDroidSlide", "Bad droid": Exit Sub
    checkDroid psDroid, "src/move.cpp:1195", "moveCalcDroidSlide", max_check_object_recursion

    bLegs = False
    If aDroid(psDroid).droidType = DROID_PERSON Or DROID_isCyborg(psDroid) Then
        bLegs = True
    End If
    spmx = gameTimeAdjustedAverage(pmx, EXTRA_PRECISION)
    spmy = gameTimeAdjustedAverage(pmy, EXTRA_PRECISION)

    droidR = moveObjRadius(psDroid)
    psObst = 0
    n = gridStartIterate(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, OBJ_MAXRADIUS, gridListSlide)
    For gi = 0 To n - 1
        psObj = gridListSlide(gi)
        If SIMPLE_OBJECT_died(psObj) <> 0 Then
            If Not (SIMPLE_OBJECT_type(psObj) < OBJ_NUM_TYPES) Then ASSERT_FAILURE "psObj->type < OBJ_NUM_TYPES", "src/move.cpp:1214", "moveCalcDroidSlide", "Bad pointer! type=%u", SIMPLE_OBJECT_type(psObj)
            GoTo continue_
        End If
        If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
            psObjcast = psObj
            objR = moveObjRadius(psObj)
            If DROID_isFlightBasedTransporter(psObjcast) Then
                ' ignore transporters
                GoTo continue_
            End If
            overFlying = False
            If Not DROID_isFlying(psDroid) And DROID_isFlying(psObjcast) Then
                overFlying = aDroid(psObjcast).pos.z > (aDroid(psDroid).pos.z + droidR)
            End If
            If Not overFlying Then
                If Not DROID_isFlying(psObjcast) And DROID_isFlying(psDroid) Then
                    overFlying = aDroid(psDroid).pos.z > (aDroid(psObjcast).pos.z + objR)
                End If
            End If
            If overFlying Then
                ' ground unit can't bump into a flying saucer..
                GoTo continue_
            End If
            If Not bLegs And aDroid(psObjcast).droidType = DROID_PERSON Then
                ' everything else doesn't avoid people
                GoTo continue_
            End If
            If aDroid(psObjcast).player = aDroid(psDroid).player _
               And aDroid(psDroid).lastFrustratedTime <> 0 _
               And ULT(gameTime - aDroid(psDroid).lastFrustratedTime, FRUSTRATED_TIME) Then
                GoTo continue_ ' clip straight through own units when sufficient frustrated -- using cheat codes!
            End If
        Else
            ' ignore anything that isn't a droid
            GoTo continue_
        End If

        objR = moveObjRadius(psObj)
        rad = droidR + objR
        radSq = rad * rad

        xdiff = aDroid(psDroid).pos.x + spmx - aDroid(psObj).pos.x
        ydiff = aDroid(psDroid).pos.y + spmy - aDroid(psObj).pos.y
        distSq = xdiff * xdiff + ydiff * ydiff
        If xdiff * spmx + ydiff * spmy >= 0 Then
            ' object behind
            GoTo continue_
        End If

        If radSq > distSq Then
            If psObst <> 0 Then
                ' hit more than one droid - stop
                pmx = 0
                pmy = 0
                psObst = 0
                Exit For
            Else
                psObst = psObj

                ' note the bump time and position if necessary
                If aDroid(psDroid).sMove.bumpTime = 0 Then
                    aDroid(psDroid).sMove.bumpTime = gameTime
                    aDroid(psDroid).sMove.lastBump = 0
                    aDroid(psDroid).sMove.pauseTime = 0
                    aDroid(psDroid).sMove.bumpPos = aDroid(psDroid).pos
                    aDroid(psDroid).sMove.bumpDir = aDroid(psDroid).rot.direction
                Else
                    aDroid(psDroid).sMove.lastBump = (gameTime - aDroid(psDroid).sMove.bumpTime) And &HFFFF&
                End If

                ' tell inactive droids to get out the way
                If psObst <> 0 Then
                    If SIMPLE_OBJECT_type(psObst) = OBJ_DROID Then
                        psShuffleDroid = psObst

                        If aiCheckAlliances(aDroid(psObst).player, aDroid(psDroid).player) _
                           And aDroid(psShuffleDroid).action <> DACTION_WAITDURINGREARM _
                           And aDroid(psShuffleDroid).sMove.Status = MOVEINACTIVE Then
                            s.x = aDroid(psDroid).sMove.target.x - aDroid(psDroid).pos.x
                            s.y = aDroid(psDroid).sMove.target.y - aDroid(psDroid).pos.y
                            moveShuffleDroid psShuffleDroid, s
                        End If
                    End If
                End If
            End If
        End If
continue_:
    Next

    If psObst <> 0 Then
        ' Try to slide round it
        moveCalcSlideVector psDroid, aDroid(psObst).pos.x, aDroid(psObst).pos.y, pmx, pmy
    End If
    checkDroid psDroid, "src/move.cpp:1313", "moveCalcDroidSlide", max_check_object_recursion
End Sub

' Get a direction for a droid to avoid obstacles etc.
' \param psDroid Which droid to examine
' \return The normalised direction vector
Private Function moveGetDirection(ByVal psDroid As Long) As Long
    Dim ctx As SteeringContext
    ' Build steering context
    ctx.droid = psDroid
    ctx.currentPos.x = aDroid(psDroid).pos.x
    ctx.currentPos.y = aDroid(psDroid).pos.y
    ctx.targetPos = aDroid(psDroid).sMove.target
    ctx.velocity.x = iSinR(aDroid(psDroid).sMove.moveDir, aDroid(psDroid).sMove.speed)
    ctx.velocity.y = iCosR(aDroid(psDroid).sMove.moveDir, aDroid(psDroid).sMove.speed)
    ctx.moveDir = aDroid(psDroid).sMove.moveDir
    ctx.speed = aDroid(psDroid).sMove.speed
    ctx.maxSpeed = asPropulsionStats(DROID_getPropulsionStats(psDroid)).maxSpeed
    ctx.radius = moveObjRadius(psDroid)

    ' Calculate steering direction
    moveSteeringManager
    moveGetDirection = SteeringManager_calculateSteeringDirection(ctx)
End Function

' Check if a droid has got to a way point
Private Function moveReachedWayPoint(ByVal psDroid As Long) As Boolean
    Dim dist As Vector2i, sqDist As Long, last As Boolean, sqprecision As Long
    Const toleranceDist As Long = 3 * (TILE_UNITS \ 4)

    ' Calculate the vector to the droid
    dist.x = aDroid(psDroid).pos.x - aDroid(psDroid).sMove.target.x
    dist.y = aDroid(psDroid).pos.y - aDroid(psDroid).sMove.target.y
    sqDist = dot2i(dist, dist)
    last = aDroid(psDroid).sMove.pathIndex = aDroid(psDroid).sMove.asPath_size
    If last Then sqprecision = (TILE_UNITS \ 4) * (TILE_UNITS \ 4) Else sqprecision = (TILE_UNITS \ 2) * (TILE_UNITS \ 2)

    ' Start the process of potentially increasing the distance threshold when very close to the waypoint.
    If ULT(sqDist, toleranceDist * toleranceDist) Then
        If aDroid(psDroid).sMove.Status <> MOVEPOINTTOPOINT Then
            aDroid(psDroid).sMove.tolerance = 0 ' If suddenly shuffling or paused, retry later.
        Else
            aDroid(psDroid).sMove.tolerance = aDroid(psDroid).sMove.tolerance + 1
        End If
    End If

    If last And aDroid(psDroid).sMove.bumpTime <> 0 Then
        ' Make waypoint tolerance 1 tile after 0 seconds, 2 tiles after 3 seconds, X tiles after (X + 1)^2 seconds.
        sqprecision = FromU32(U32(gameTime - aDroid(psDroid).sMove.bumpTime + GAME_TICKS_PER_SEC) * (TILE_UNITS * TILE_UNITS \ GAME_TICKS_PER_SEC))
    ElseIf Not ULT(aDroid(psDroid).sMove.tolerance, toleranceDist) Then
        sqprecision = MUL32(aDroid(psDroid).sMove.tolerance, aDroid(psDroid).sMove.tolerance)
    End If

    ' Else check current waypoint
    If ULT(sqDist, sqprecision) Then
        aDroid(psDroid).sMove.tolerance = 0 ' Reset threshold tolerance counter when successful.
        moveReachedWayPoint = True
        Exit Function
    End If
    moveReachedWayPoint = False
End Function

' Calculate the new speed for a droid based on factors like pitch.
' @todo Remove hack for steep slopes not properly marked as blocking on some maps.
Public Function moveCalcDroidSpeed(ByVal psDroid As Long) As Long
    Dim maxPitch As Long, mapX As Long, mapY As Long, speed As Long, pitch As Long, psWStats As Long, FrmSpeed As Long

    maxPitch = DEG(MAX_SPEED_PITCH)

    checkDroid psDroid, "src/move.cpp:1392", "moveCalcDroidSpeed", max_check_object_recursion

    ' NOTE: This screws up since the transporter is offscreen still (on a mission!), and we are trying to find terrainType of a tile (that is offscreen!)
    If aDroid(psDroid).droidType = DROID_SUPERTRANSPORTER And missionIsOffworld() Then
        speed = asPropulsionStats(DROID_getPropulsionStats(psDroid)).maxSpeed
    Else
        mapX = map_coord(aDroid(psDroid).pos.x)
        mapY = map_coord(aDroid(psDroid).pos.y)
        speed = calcDroidSpeed(aDroid(psDroid).baseSpeed, terrainType(mapTile(mapX, mapY)), aDroid(psDroid).asBits(COMP_PROPULSION), getDroidEffectiveLevel(psDroid, False))
    End If

    ' now offset the speed for the slope of the droid
    pitch = angleDelta(aDroid(psDroid).rot.pitch)
    speed = (maxPitch - pitch) * speed \ maxPitch
    If speed <= 10 Then
        ' Very nasty hack to deal with buggy maps, where some cliffs are
        ' not properly marked as being cliffs, but too steep to drive over.
        ' This confuses the heck out of the path-finding code! - Per
        speed = 10
    End If

    ' stop droids that have just fired a no fire while moving weapon
    If aDroid(psDroid).numWeaps > 0 Then
        If aDroid(psDroid).asWeaps(0).nStat <> 0 And ULT(gameTime, aDroid(psDroid).asWeaps(0).lastFired + FOM_MOVEPAUSE) Then
            psWStats = BASE_OBJECT_getWeaponStats(psDroid, 0)
            If Not asWeaponStats(psWStats).fireOnMove Then
                speed = 0
            End If
        End If
    End If

    ' adjust speed for formation
    If Not DROID_isVtol(psDroid) _
       And moveFormationSpeedLimitingOn(aDroid(psDroid).player) And aDroid(psDroid).sMove.psFormation <> 0 _
       And Not DROID_isRetreatingForRepair(psDroid) And Not DROID_isReturningToBase(psDroid) Then
        FrmSpeed = aFormation(aDroid(psDroid).sMove.psFormation).iSpeed

        If speed > FrmSpeed Then
            speed = FrmSpeed
        End If
    End If

    ' slow down shuffling VTOLs
    If DROID_isVtol(psDroid) And _
       (aDroid(psDroid).sMove.Status = MOVESHUFFLE) And _
       (speed > MIN_END_SPEED) Then
        speed = MIN_END_SPEED
    End If

    checkDroid psDroid, "src/move.cpp:1453", "moveCalcDroidSpeed", max_check_object_recursion

    moveCalcDroidSpeed = speed
End Function

' Determine whether a droid has stopped moving.
' @return true if the droid doesn't move, false if it's moving.
Private Function moveDroidStopped(ByVal psDroid As Long, ByVal speed As Long) As Boolean
    If aDroid(psDroid).sMove.Status = MOVEINACTIVE And speed = 0 And aDroid(psDroid).sMove.speed = 0 Then
        moveDroidStopped = True
    Else
        moveDroidStopped = False
    End If
End Function

' Direction is target direction.
Private Sub moveUpdateDroidDirection(ByVal psDroid As Long, pSpeed As Long, ByVal direction As Long, _
                                     ByVal iSpinAngle As Long, ByVal iSpinSpeed As Long, ByVal iTurnSpeed As Long, pDroidDir As Long)
    Dim diff As Long, turnSpeed As Long, maxChange As Long

    iSpinAngle = iSpinAngle And &HFFFF&     ' uint16_t
    pDroidDir = aDroid(psDroid).rot.direction

    ' don't move if in MOVEPAUSE state
    If aDroid(psDroid).sMove.Status = MOVEPAUSE Then
        Exit Sub
    End If

    diff = angleDelta(direction - pDroidDir)
    ' Turn while moving - slow down speed depending on target angle so that we can turn faster
    pSpeed = MAX_(pSpeed * (iSpinAngle - Abs(diff)) \ iSpinAngle, 0)

    ' iTurnSpeed is turn speed at max velocity, increase turn speed up to iSpinSpeed when slowing down
    turnSpeed = MIN_(FromU32(iTurnSpeed + Fix(CDbl(iSpinSpeed - iTurnSpeed) * Abs(diff) / iSpinAngle)), iSpinSpeed)

    ' Calculate the maximum change in direction
    maxChange = gameTimeAdjustedAverage(turnSpeed)

    ' Move *pDroidDir towards target, by at most maxChange.
    pDroidDir = (pDroidDir + clip_(diff, -maxChange, maxChange)) And &HFFFF&
End Sub

' Calculate current speed perpendicular to droids direction
Private Function moveCalcPerpSpeed(ByVal psDroid As Long, ByVal iDroidDir As Long, ByVal iSkidDecel As Long) As Long
    Dim adiff As Long, perpSpeed As Long
    adiff = angleDelta(iDroidDir - aDroid(psDroid).sMove.moveDir)
    perpSpeed = iSinR(Abs(adiff), aDroid(psDroid).sMove.speed)

    ' decelerate the perpendicular speed
    perpSpeed = MAX_(0, perpSpeed - gameTimeAdjustedAverage(iSkidDecel))

    moveCalcPerpSpeed = perpSpeed
End Function

Private Sub moveCombineNormalAndPerpSpeeds(ByVal psDroid As Long, ByVal fNormalSpeed As Long, ByVal fPerpSpeed As Long, ByVal iDroidDir As Long)
    Dim adiff As Long, relDir As Long, finalSpeed As Long

    ' set current direction
    aDroid(psDroid).rot.direction = iDroidDir

    ' set normal speed and direction if perpendicular speed is zero
    If fPerpSpeed = 0 Then
        aDroid(psDroid).sMove.speed = fNormalSpeed
        aDroid(psDroid).sMove.moveDir = iDroidDir
        Exit Sub
    End If

    finalSpeed = iHypot(fNormalSpeed, fPerpSpeed)

    ' calculate the angle between the droid facing and movement direction
    relDir = iAtan2(fPerpSpeed, fNormalSpeed)

    ' choose the finalDir on the same side as the old movement direction
    adiff = angleDelta(iDroidDir - aDroid(psDroid).sMove.moveDir)

    If adiff < 0 Then    ' Cast wrapping intended.
        aDroid(psDroid).sMove.moveDir = (iDroidDir + relDir) And &HFFFF&
    Else
        aDroid(psDroid).sMove.moveDir = (iDroidDir - relDir) And &HFFFF&
    End If
    aDroid(psDroid).sMove.speed = finalSpeed
End Sub

' Calculate the current speed in the droids normal direction
Private Function moveCalcNormalSpeed(ByVal psDroid As Long, ByVal fSpeed As Long, ByVal iDroidDir As Long, ByVal iAccel As Long, ByVal iDecel As Long) As Long
    Dim adiff As Long, normalSpeed As Long

    adiff = (iDroidDir - aDroid(psDroid).sMove.moveDir) And &HFFFF&  ' Cast wrapping intended.
    normalSpeed = iCosR(adiff, aDroid(psDroid).sMove.speed)

    If normalSpeed < fSpeed Then
        ' accelerate
        normalSpeed = normalSpeed + gameTimeAdjustedAverage(iAccel)
        If normalSpeed > fSpeed Then
            normalSpeed = fSpeed
        End If
    Else
        ' decelerate
        normalSpeed = normalSpeed - gameTimeAdjustedAverage(iDecel)
        If normalSpeed < fSpeed Then
            normalSpeed = fSpeed
        End If
    End If

    moveCalcNormalSpeed = normalSpeed
End Function

Private Sub moveGetDroidPosDiffs(ByVal psDroid As Long, pDX As Long, pDY As Long)
    Dim move As Long
    move = aDroid(psDroid).sMove.speed * EXTRA_PRECISION  ' high precision

    pDX = iSinR(aDroid(psDroid).sMove.moveDir, move)
    pDY = iCosR(aDroid(psDroid).sMove.moveDir, move)
End Sub

' see if the droid is close to the final way point
Private Sub moveCheckFinalWaypoint(ByVal psDroid As Long, pSpeed As Long)
    Dim minEndSpeed As Long, diff As Vector2i, distSq As Long
    minEndSpeed = (pSpeed + 2) \ 3
    minEndSpeed = MIN_(minEndSpeed, MIN_END_SPEED)

    ' don't do this for VTOLs doing attack runs
    If DROID_isVtol(psDroid) And (aDroid(psDroid).action = DACTION_VTOLATTACK) Then
        Exit Sub
    End If

    If aDroid(psDroid).sMove.Status <> MOVESHUFFLE And _
       aDroid(psDroid).sMove.pathIndex = aDroid(psDroid).sMove.asPath_size Then
        diff.x = aDroid(psDroid).pos.x - aDroid(psDroid).sMove.target.x
        diff.y = aDroid(psDroid).pos.y - aDroid(psDroid).sMove.target.y
        distSq = dot2i(diff, diff)
        If distSq < END_SPEED_RANGE * END_SPEED_RANGE Then
            pSpeed = (pSpeed - minEndSpeed) * distSq \ (END_SPEED_RANGE * END_SPEED_RANGE) + minEndSpeed
        End If
    End If
End Sub

Private Sub moveUpdateDroidPos(ByVal psDroid As Long, ByVal dx As Long, ByVal dy As Long)
    checkDroid psDroid, "src/move.cpp:1608", "moveUpdateDroidPos", max_check_object_recursion

    If aDroid(psDroid).sMove.Status = MOVEPAUSE Or isDead(psDroid) Then
        ' don't actually move if the move is paused
        Exit Sub
    End If

    aDroid(psDroid).pos.x = aDroid(psDroid).pos.x + gameTimeAdjustedAverage(dx, EXTRA_PRECISION)
    aDroid(psDroid).pos.y = aDroid(psDroid).pos.y + gameTimeAdjustedAverage(dy, EXTRA_PRECISION)

    ' impact if about to go off map else update coordinates
    If worldOnMap(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y) = False Then
        ' transporter going off-world will trigger next map, and is ok
        If Not (DROID_isTransporter(psDroid)) Then ASSERT_FAILURE "psDroid->isTransporter()", "src/move.cpp:1623", "moveUpdateDroidPos", "droid trying to move off the map!"
        If Not DROID_isTransporter(psDroid) Then
            ' dreadful last-ditch crash-avoiding hack - sort this! - GJ
            destroyDroid psDroid, gameTime
            Exit Sub
        End If
    End If

    ' lovely hack to keep transporters just on the map
    ' two weeks to go and the hacks just get better !!!
    If DROID_isTransporter(psDroid) Then
        If aDroid(psDroid).pos.x = 0 Then
            aDroid(psDroid).pos.x = 1
        End If
        If aDroid(psDroid).pos.y = 0 Then
            aDroid(psDroid).pos.y = 1
        End If
    End If
    checkDroid psDroid, "src/move.cpp:1645", "moveUpdateDroidPos", max_check_object_recursion
End Sub

' Update a tracked droids position and speed given target values
Private Sub moveUpdateGroundModel(ByVal psDroid As Long, ByVal speed As Long, ByVal direction As Long)
    Dim fPerpSpeed As Long, fNormalSpeed As Long, iDroidDir As Long, slideDir As Long, psPropStats As Long
    Dim spinSpeed As Long, spinAngle As Long, turnSpeed As Long, dx As Long, dy As Long, bx As Long, by As Long

    checkDroid psDroid, "src/move.cpp:1657", "moveUpdateGroundModel", max_check_object_recursion

    ' nothing to do if the droid is stopped
    If moveDroidStopped(psDroid, speed) = True Then
        Exit Sub
    End If

    psPropStats = DROID_getPropulsionStats(psDroid)
    spinSpeed = aDroid(psDroid).baseSpeed * asPropulsionStats(psPropStats).spinSpeed
    turnSpeed = aDroid(psDroid).baseSpeed * asPropulsionStats(psPropStats).turnSpeed
    spinAngle = DEG(asPropulsionStats(psPropStats).spinAngle)

    moveCheckFinalWaypoint psDroid, speed

    moveUpdateDroidDirection psDroid, speed, direction, spinAngle, spinSpeed, turnSpeed, iDroidDir

    fNormalSpeed = moveCalcNormalSpeed(psDroid, speed, iDroidDir, asPropulsionStats(psPropStats).acceleration, asPropulsionStats(psPropStats).deceleration)
    fPerpSpeed = moveCalcPerpSpeed(psDroid, iDroidDir, asPropulsionStats(psPropStats).skidDeceleration)

    moveCombineNormalAndPerpSpeeds psDroid, fNormalSpeed, fPerpSpeed, iDroidDir
    moveGetDroidPosDiffs psDroid, dx, dy
    moveOpenGates psDroid
    moveCheckSquished psDroid, dx, dy
    moveCalcDroidSlide psDroid, dx, dy
    bx = dx
    by = dy
    moveCalcBlockingSlide psDroid, bx, by, direction, slideDir
    If bx <> dx Or by <> dy Then
        moveUpdateDroidDirection psDroid, speed, slideDir, spinAngle, aDroid(psDroid).baseSpeed * DEG(1), aDroid(psDroid).baseSpeed * DEG(1) \ 3, iDroidDir
        aDroid(psDroid).rot.direction = iDroidDir
    End If

    moveUpdateDroidPos psDroid, bx, by

    'set the droid height here so other routines can use it
    aDroid(psDroid).pos.z = map_Height(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y) 'jps 21july96
    updateDroidOrientation psDroid
End Sub

' Update a persons position and speed given target values
Private Sub moveUpdatePersonModel(ByVal psDroid As Long, ByVal speed As Long, ByVal direction As Long)
    Dim fPerpSpeed As Long, fNormalSpeed As Long, spinSpeed As Long, turnSpeed As Long, dx As Long, dy As Long
    Dim iDroidDir As Long, slideDir As Long, psPropStats As Long

    checkDroid psDroid, "src/move.cpp:1707", "moveUpdatePersonModel", max_check_object_recursion

    ' if the droid is stopped, only make sure animations are set correctly
    If moveDroidStopped(psDroid, speed) Then
        If aDroid(psDroid).droidType = DROID_PERSON And _
           (aDroid(psDroid).action = DACTION_ATTACK Or _
           aDroid(psDroid).action = DACTION_ROTATETOATTACK) _
           And aDroid(psDroid).animationEvent <> ANIM_EVENT_DYING _
           And aDroid(psDroid).animationEvent <> ANIM_EVENT_FIRING Then
            aDroid(psDroid).timeAnimationStarted = gameTime
            aDroid(psDroid).animationEvent = ANIM_EVENT_FIRING
        ElseIf aDroid(psDroid).animationEvent = ANIM_EVENT_ACTIVE Then
            resetObjectAnimationState psDroid
        End If
        Exit Sub
    End If

    psPropStats = DROID_getPropulsionStats(psDroid)
    spinSpeed = aDroid(psDroid).baseSpeed * asPropulsionStats(psPropStats).spinSpeed
    turnSpeed = aDroid(psDroid).baseSpeed * asPropulsionStats(psPropStats).turnSpeed

    moveUpdateDroidDirection psDroid, speed, direction, DEG(asPropulsionStats(psPropStats).spinAngle), spinSpeed, turnSpeed, iDroidDir

    fNormalSpeed = moveCalcNormalSpeed(psDroid, speed, iDroidDir, asPropulsionStats(psPropStats).acceleration, asPropulsionStats(psPropStats).deceleration)

    ' people don't skid at the moment so set zero perpendicular speed
    fPerpSpeed = 0

    moveCombineNormalAndPerpSpeeds psDroid, fNormalSpeed, fPerpSpeed, iDroidDir
    moveGetDroidPosDiffs psDroid, dx, dy
    moveOpenGates psDroid
    moveCalcDroidSlide psDroid, dx, dy
    moveCalcBlockingSlide psDroid, dx, dy, direction, slideDir
    moveUpdateDroidPos psDroid, dx, dy

    'set the droid height here so other routines can use it
    aDroid(psDroid).pos.z = map_Height(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y) 'jps 21july96

    ' update anim if moving
    If aDroid(psDroid).droidType = DROID_PERSON And speed <> 0 And (aDroid(psDroid).animationEvent <> ANIM_EVENT_ACTIVE And aDroid(psDroid).animationEvent <> ANIM_EVENT_DYING) Then
        aDroid(psDroid).timeAnimationStarted = gameTime
        aDroid(psDroid).animationEvent = ANIM_EVENT_ACTIVE
    End If

    checkDroid psDroid, "src/move.cpp:1756", "moveUpdatePersonModel", max_check_object_recursion
End Sub

' VTOL_VERTICAL_SPEED_OLD
Private Function VTOL_VERTICAL_SPEED_OLD(ByVal psDroid As Long) As Long
    If aDroid(psDroid).baseSpeed \ 4 > 60 Then VTOL_VERTICAL_SPEED_OLD = aDroid(psDroid).baseSpeed \ 4 Else VTOL_VERTICAL_SPEED_OLD = 60
End Function

'Slower VTOLs need to increase height faster than the above or else they hover too close to the ground
'during flight paths that feature increasing terrain height. Resulting in poor attack angles in some cases.
Private Function VTOL_VERTICAL_SPEED(ByVal psDroid As Long) As Long
    If aDroid(psDroid).baseSpeed \ 4 > 160 Then VTOL_VERTICAL_SPEED = aDroid(psDroid).baseSpeed \ 4 Else VTOL_VERTICAL_SPEED = 160
End Function

' primitive 'bang-bang' vtol height controller
Private Sub moveAdjustVtolHeight(ByVal psDroid As Long, ByVal iMapHeight As Long)
    Dim iMinHeight As Long, iMaxHeight As Long, iLevelHeight As Long
    If DROID_isTransporter(psDroid) And Not bMultiPlayer Then
        iMinHeight = 2 * VTOL_HEIGHT_MIN
        iLevelHeight = 2 * VTOL_HEIGHT_LEVEL
        iMaxHeight = 2 * VTOL_HEIGHT_MAX
    Else
        iMinHeight = VTOL_HEIGHT_MIN
        iLevelHeight = VTOL_HEIGHT_LEVEL
        iMaxHeight = VTOL_HEIGHT_MAX
    End If

    If aDroid(psDroid).pos.z >= (iMapHeight + iMaxHeight) Then
        aDroid(psDroid).sMove.iVertSpeed = ToInt16(-VTOL_VERTICAL_SPEED_OLD(psDroid))
    ElseIf aDroid(psDroid).pos.z < (iMapHeight + iMinHeight) Then
        If DROID_isTransporter(psDroid) Then
            aDroid(psDroid).sMove.iVertSpeed = ToInt16(VTOL_VERTICAL_SPEED_OLD(psDroid))
        Else
            aDroid(psDroid).sMove.iVertSpeed = ToInt16(VTOL_VERTICAL_SPEED(psDroid))
        End If
    ElseIf (aDroid(psDroid).pos.z < iLevelHeight) And _
           (aDroid(psDroid).sMove.iVertSpeed < 0) Then
        aDroid(psDroid).sMove.iVertSpeed = 0
    ElseIf (aDroid(psDroid).pos.z > iLevelHeight) And _
           (aDroid(psDroid).sMove.iVertSpeed > 0) Then
        aDroid(psDroid).sMove.iVertSpeed = 0
    End If
End Sub

Private Sub moveUpdateVtolModel(ByVal psDroid As Long, ByVal speed As Long, ByVal direction As Long)
    Dim fPerpSpeed As Long, fNormalSpeed As Long, iDroidDir As Long, slideDir As Long
    Dim spinSpeed As Long, turnSpeed As Long, iMapZ As Long, iSpinSpeed As Long, iTurnSpeed As Long, dx As Long, dy As Long
    Dim targetRoll As Long, psPropStats As Long

    checkDroid psDroid, "src/move.cpp:1817", "moveUpdateVtolModel", max_check_object_recursion

    ' nothing to do if the droid is stopped
    If moveDroidStopped(psDroid, speed) = True Then
        Exit Sub
    End If

    psPropStats = DROID_getPropulsionStats(psDroid)
    spinSpeed = DEG(asPropulsionStats(psPropStats).spinSpeed)
    turnSpeed = DEG(asPropulsionStats(psPropStats).turnSpeed)

    moveCheckFinalWaypoint psDroid, speed

    If DROID_isTransporter(psDroid) Then
        moveUpdateDroidDirection psDroid, speed, direction, DEG(asPropulsionStats(psPropStats).spinAngle), spinSpeed, turnSpeed, iDroidDir
    Else
        iSpinSpeed = MAX_(aDroid(psDroid).baseSpeed * DEG(1) \ 2, spinSpeed)
        iTurnSpeed = MAX_(aDroid(psDroid).baseSpeed * DEG(1) \ 8, turnSpeed)
        moveUpdateDroidDirection psDroid, speed, direction, DEG(asPropulsionStats(psPropStats).spinAngle), iSpinSpeed, iTurnSpeed, iDroidDir
    End If

    fNormalSpeed = moveCalcNormalSpeed(psDroid, speed, iDroidDir, asPropulsionStats(psPropStats).acceleration, asPropulsionStats(psPropStats).deceleration)
    fPerpSpeed = moveCalcPerpSpeed(psDroid, iDroidDir, asPropulsionStats(psPropStats).skidDeceleration)

    moveCombineNormalAndPerpSpeeds psDroid, fNormalSpeed, fPerpSpeed, iDroidDir

    moveGetDroidPosDiffs psDroid, dx, dy

    ' set slide blocking tile for map edge
    If Not DROID_isTransporter(psDroid) Then
        moveCalcBlockingSlide psDroid, dx, dy, direction, slideDir
    End If

    moveUpdateDroidPos psDroid, dx, dy

    ' update vtol orientation
    targetRoll = clip_(4 * angleDelta(aDroid(psDroid).sMove.moveDir - aDroid(psDroid).rot.direction), -DEG(60), DEG(60)) And &HFFFF&
    aDroid(psDroid).rot.roll = (aDroid(psDroid).rot.roll + (gameTimeAdjustedIncrement(3 * angleDelta(targetRoll - aDroid(psDroid).rot.roll)) And &HFFFF&)) And &HFFFF&

    ' do vertical movement - only if on the map
    If worldOnMap(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y) Then
        iMapZ = map_Height(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y)
        aDroid(psDroid).pos.z = MAX_(iMapZ, aDroid(psDroid).pos.z + gameTimeAdjustedIncrement(aDroid(psDroid).sMove.iVertSpeed))
        moveAdjustVtolHeight psDroid, iMapZ
        aDroid(psDroid).heightAboveMap = aDroid(psDroid).pos.z - iMapZ
    End If
End Sub

Private Sub moveUpdateCyborgModel(ByVal psDroid As Long, ByVal moveSpeed As Long, ByVal moveDir As Long, ByVal oldStatus As Long)
    checkDroid psDroid, "src/move.cpp:1873", "moveUpdateCyborgModel", max_check_object_recursion

    ' nothing to do if the droid is stopped
    If moveDroidStopped(psDroid, moveSpeed) Then
        If aDroid(psDroid).animationEvent = ANIM_EVENT_ACTIVE Then
            resetObjectAnimationState psDroid
        End If
        Exit Sub
    End If

    If aDroid(psDroid).animationEvent = ANIM_EVENT_NONE Then
        aDroid(psDroid).timeAnimationStarted = gameTime
        aDroid(psDroid).animationEvent = ANIM_EVENT_ACTIVE
    End If

    ' use baba person movement
    moveUpdatePersonModel psDroid, moveSpeed, moveDir

    aDroid(psDroid).rot.pitch = 0
    aDroid(psDroid).rot.roll = 0
End Sub

Private Sub moveDescending(ByVal psDroid As Long)
    Dim iMapHeight As Long
    iMapHeight = map_Height(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y)

    aDroid(psDroid).sMove.speed = 0

    If aDroid(psDroid).pos.z > iMapHeight Then
        ' descending
        aDroid(psDroid).sMove.iVertSpeed = ToInt16(-VTOL_VERTICAL_SPEED_OLD(psDroid))
    Else
        ' on floor - stop
        aDroid(psDroid).pos.z = iMapHeight
        aDroid(psDroid).sMove.iVertSpeed = 0

        ' reset move state
        aDroid(psDroid).sMove.Status = MOVEINACTIVE

        ' conform to terrain
        updateDroidOrientation psDroid
    End If
End Sub

' audio callback used to kill movement sounds
Public Function moveCheckDroidMovingAndVisible(ByVal psObj As Long) As Boolean
    Dim psDroid As Long
    psDroid = psObj

    If psDroid = 0 Then
        moveCheckDroidMovingAndVisible = False
        Exit Function
    End If

    ' check for dead, not moving or invisible to player
    If aDroid(psDroid).died <> 0 Or moveDroidStopped(psDroid, 0) Or _
       (DROID_isTransporter(psDroid) And aDroid(psDroid).order.typ = DORDER_NONE) Or _
       Not (BASE_OBJECT_visibleForLocalDisplay(psDroid) <> 0) Then
        aDroid(psDroid).iAudioID = NO_SOUND
        moveCheckDroidMovingAndVisible = False
        Exit Function
    End If

    moveCheckDroidMovingAndVisible = True
End Function

Private Sub movePlayDroidMoveAudio(ByVal psDroid As Long)
    Dim iAudioID As Long, iPropType As Long, psPropStats As Long

    iAudioID = NO_SOUND
    iPropType = 0

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/move.cpp:1952", "movePlayDroidMoveAudio", "Unit pointer invalid": Exit Sub

    If (psDroid <> 0) Then
        If BASE_OBJECT_visibleForLocalDisplay(psDroid) <> 0 Then
            psPropStats = DROID_getPropulsionStats(psDroid)
            If Not (psPropStats <> -1) Then ASSERT_FAILURE "psPropStats != nullptr", "src/move.cpp:1958", "movePlayDroidMoveAudio", "Invalid propulsion stats pointer": Exit Sub
            iPropType = asPropulsionStats(psPropStats).propulsionType And &HFF&

            ' play specific wheeled and transporter or stats-specified noises
            If iPropType = PROPULSION_TYPE_WHEELED And aDroid(psDroid).droidType <> DROID_CONSTRUCT Then
                iAudioID = ID_SOUND_TREAD
            ElseIf DROID_isTransporter(psDroid) Then
                iAudioID = ID_SOUND_BLIMP_FLIGHT
            ElseIf iPropType = PROPULSION_TYPE_LEGGED And DROID_isCyborg(psDroid) And asPropulsionTypes(iPropType).moveID = NO_SOUND Then
                iAudioID = ID_SOUND_CYBORG_MOVE
            Else
                iAudioID = asPropulsionTypes(iPropType).moveID
            End If

            If iAudioID <> NO_SOUND Then
                If audio_PlayObjDynamicTrack(psDroid, iAudioID, _
                                             CB_moveCheckDroidMovingAndVisible) Then
                    aDroid(psDroid).iAudioID = iAudioID
                End If
            End If
        End If
    End If
End Sub

' audio callback: the start sound has finished
Public Function moveDroidStartCallback(ByVal psObj As Long) As Boolean
    Dim psDroid As Long
    psDroid = psObj

    If psDroid = 0 Then
        moveDroidStartCallback = False
        Exit Function
    End If

    movePlayDroidMoveAudio psDroid

    moveDroidStartCallback = True
End Function

Private Sub movePlayAudio(ByVal psDroid As Long, ByVal bStarted As Boolean, ByVal bStoppedBefore As Boolean, ByVal iMoveSpeed As Long)
    Dim propType As Long, psPropStats As Long, bStoppedNow As Boolean, iAudioID As Long, pAudioCallback As Long

    iAudioID = NO_SOUND
    pAudioCallback = 0

    ' get prop stats
    psPropStats = DROID_getPropulsionStats(psDroid)
    If Not (psPropStats <> -1) Then ASSERT_FAILURE "psPropStats != nullptr", "src/move.cpp:2018", "movePlayAudio", "Invalid propulsion stats pointer": Exit Sub
    propType = asPropulsionStats(psPropStats).propulsionType And &HFF&

    ' get current droid motion status
    bStoppedNow = moveDroidStopped(psDroid, iMoveSpeed)

    If bStarted Then
        ' play start audio
        If (propType = PROPULSION_TYPE_WHEELED And aDroid(psDroid).droidType <> DROID_CONSTRUCT) _
           Or asPropulsionTypes(propType).startID = NO_SOUND Then
            movePlayDroidMoveAudio psDroid
            Exit Sub
        ElseIf DROID_isTransporter(psDroid) Then
            iAudioID = ID_SOUND_BLIMP_TAKE_OFF
        Else
            iAudioID = asPropulsionTypes(propType).startID
        End If

        pAudioCallback = CB_moveDroidStartCallback
    ElseIf Not bStoppedBefore And bStoppedNow And _
           (asPropulsionTypes(propType).shutDownID <> NO_SOUND) Then
        ' play stop audio
        If DROID_isTransporter(psDroid) Then
            iAudioID = ID_SOUND_BLIMP_LAND
        ElseIf propType <> PROPULSION_TYPE_WHEELED Or aDroid(psDroid).droidType = DROID_CONSTRUCT Then
            iAudioID = asPropulsionTypes(propType).shutDownID
        End If
    ElseIf Not bStoppedBefore And Not bStoppedNow And aDroid(psDroid).iAudioID = NO_SOUND Then
        ' play move audio
        movePlayDroidMoveAudio psDroid
        Exit Sub
    End If

    If (iAudioID <> NO_SOUND) And _
       (BASE_OBJECT_visibleForLocalDisplay(psDroid) <> 0) Then
        If audio_PlayObjDynamicTrack(psDroid, iAudioID, _
                                     pAudioCallback) Then
            aDroid(psDroid).iAudioID = iAudioID
        End If
    End If
End Sub

Private Function pickupOilDrum(ByVal toPlayer As Long, ByVal fromPlayer As Long) As Boolean
    Dim power As Long
    power = OILDRUM_POWER

    If Not bMultiPlayer And Not bInTutorial Then
        ' Let Beta and Gamma campaign oil drums give a little more power
        If getCampaignNumber() = 2 Then
            power = OILDRUM_POWER + (OILDRUM_POWER \ 2)
        ElseIf getCampaignNumber() = 3 Then
            power = OILDRUM_POWER * 2
        End If
    End If

    addPower toPlayer, power  ' give power

    If toPlayer = selectedPlayer Then
        CONPRINTF "You found %u power in an oil drum.", power
    End If

    pickupOilDrum = True
End Function

' called when a droid moves to a new tile.
' use to pick up oil, etc..
Private Sub checkLocalFeatures(ByVal psDroid As Long)
    Dim n As Long, gi As Long, psObj As Long, pickedUp As Boolean

    ' NOTE: Why not do this for AI units also?
    If (Not isHumanPlayer(aDroid(psDroid).player) And aDroid(psDroid).order.typ <> DORDER_RECOVER) Or DROID_isVtol(psDroid) Or DROID_isTransporter(psDroid) Then  ' VTOLs or transporters can't pick up features!
        Exit Sub
    End If

    ' scan the neighbours
    n = gridStartIterate(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, DROIDDIST, gridListLocal)
    For gi = 0 To n - 1
        psObj = gridListLocal(gi)
        pickedUp = False

        If SIMPLE_OBJECT_type(psObj) = OBJ_FEATURE And SIMPLE_OBJECT_died(psObj) = 0 Then
            If Abs(aDroid(psDroid).pos.z - aFeature(psObj).pos.z) > MAX_PICKUP_DISTANCE Then
                GoTo continue_ ' Don't pickup if height difference is more than half of a tile.
            End If

            Select Case asFeatureStats(aFeature(psObj).psStats).subType
            Case FEAT_OIL_DRUM
                pickedUp = pickupOilDrum(aDroid(psDroid).player, aFeature(psObj).player)
                triggerEventPickup psObj, psDroid
            Case FEAT_GEN_ARTE
                pickedUp = pickupArtefact(aDroid(psDroid).player, aFeature(psObj).player)
                triggerEventPickup psObj, psDroid
            Case Else
            End Select
        End If

        If Not pickedUp Then
            ' Object is not a living oil drum or artefact.
            GoTo continue_
        End If

        turnOffMultiMsg True
        removeFeature psObj  ' remove artifact+.
        turnOffMultiMsg False
continue_:
    Next
End Sub

' Frame update for the movement of a tracked droid
Public Sub moveUpdateDroid(ByVal psDroid As Long)
    Dim oldx As Long, oldy As Long, oldStatus As Long, moveSpeed As Long, moveDir As Long, psPropStats As Long
    Dim pos As Vector3i, bStarted As Boolean, bStopped As Boolean, status As Long
    Dim doNavigate As Boolean, doPointToPoint As Boolean, tar As Vector2i, ok As Boolean, turning As Boolean

    oldStatus = aDroid(psDroid).sMove.Status And &HFF&
    bStarted = False

    checkDroid psDroid, "src/move.cpp:2170", "moveUpdateDroid", max_check_object_recursion

    psPropStats = DROID_getPropulsionStats(psDroid)
    If Not (psPropStats <> -1) Then ASSERT_FAILURE "psPropStats != nullptr", "src/move.cpp:2173", "moveUpdateDroid", "Invalid propulsion stats pointer": Exit Sub

    ' If the droid has been attacked by an EMP weapon, it is temporarily disabled
    If aDroid(psDroid).lastHitWeapon = WSC_EMP Then
        If ULT(gameTime - aDroid(psDroid).timeLastHit, EMP_DISABLE_TIME) Then
            ' Get out without updating
            Exit Sub
        End If
    End If

    ' save current motion status of droid
    bStopped = moveDroidStopped(psDroid, 0)

    moveSpeed = 0
    moveDir = aDroid(psDroid).rot.direction

    ' (VB) the switch, with its fallthroughs written as flags
    status = aDroid(psDroid).sMove.Status
    If status = MOVEINACTIVE Then
        If aDroid(psDroid).animationEvent = ANIM_EVENT_ACTIVE Then
            resetObjectAnimationState psDroid
        End If
        If bStopped Then
            If aDroid(psDroid).sMove.psFormation <> 0 And (aDroid(psDroid).sMove.pathIndex = aDroid(psDroid).sMove.asPath_size) Then
                ' Remove from formation - NOTE: This differs from the original formation logic
                formationLeave aDroid(psDroid).sMove.psFormation, psDroid
                aDroid(psDroid).sMove.psFormation = 0
            End If
        End If
    ElseIf status = MOVESHUFFLE Then
        If moveReachedWayPoint(psDroid) Or ULT(aDroid(psDroid).sMove.shuffleStart + MOVE_SHUFFLETIME, gameTime) Then
            If asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_LIFT Then
                aDroid(psDroid).sMove.Status = MOVEHOVER
            Else
                aDroid(psDroid).sMove.Status = MOVEINACTIVE
            End If
        Else
            ' Calculate a target vector
            moveDir = moveGetDirection(psDroid)

            moveSpeed = moveCalcDroidSpeed(psDroid)
        End If
    ElseIf status = MOVEWAITROUTE Or status = MOVENAVIGATE Or status = MOVEPOINTTOPOINT Or status = MOVEPAUSE Then
        doNavigate = status = MOVENAVIGATE
        doPointToPoint = True
        If status = MOVEWAITROUTE Then
            moveDroidTo psDroid, aDroid(psDroid).sMove.destination.x, aDroid(psDroid).sMove.destination.y
            moveSpeed = MAX_(0, aDroid(psDroid).sMove.speed - 1)
            If aDroid(psDroid).sMove.Status <> MOVENAVIGATE Then
                doPointToPoint = False
            Else
                doNavigate = True   ' fallthrough
            End If
        End If
        If doNavigate Then
            ' Get the next control point
            If Not moveNextTarget(psDroid) Then
                ' No more waypoints - finish
                If asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_LIFT Then
                    aDroid(psDroid).sMove.Status = MOVEHOVER
                Else
                    aDroid(psDroid).sMove.Status = MOVEINACTIVE
                End If
                doPointToPoint = False
            Else
                If DROID_isVtol(psDroid) Then
                    aDroid(psDroid).rot.pitch = 0
                End If

                aDroid(psDroid).sMove.Status = MOVEPOINTTOPOINT
                aDroid(psDroid).sMove.bumpTime = 0
                moveSpeed = MAX_(0, aDroid(psDroid).sMove.speed - 1)

                ' save started status for movePlayAudio
                If aDroid(psDroid).sMove.speed = 0 Then
                    bStarted = True
                End If
                ' fallthrough
            End If
        End If
        If doPointToPoint Then
            ' moving between two way points
            If aDroid(psDroid).sMove.asPath_size = 0 Then
                debug_ 2272, LOG_WARNING, "moveUpdateDroid", "No path to follow, but psDroid->sMove.Status = %d", aDroid(psDroid).sMove.Status
            End If

            ' Get the best control point.
            If aDroid(psDroid).sMove.asPath_size = 0 Then
                ' Got stuck somewhere, can't find the path.
                moveDroidTo psDroid, aDroid(psDroid).sMove.destination.x, aDroid(psDroid).sMove.destination.y
            ElseIf Not moveBestTarget(psDroid) Then
                moveDroidTo psDroid, aDroid(psDroid).sMove.destination.x, aDroid(psDroid).sMove.destination.y
            End If

            ' See if the target point has been reached
            If moveReachedWayPoint(psDroid) Then
                ' Got there - move onto the next waypoint
                If Not moveNextTarget(psDroid) Then
                    ' No more waypoints - finish
                    If asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_LIFT Then
                        ' check the location for vtols
                        tar.x = aDroid(psDroid).pos.x
                        tar.y = aDroid(psDroid).pos.y
                        ' Not doing an order which means we never land (which means we might want to land).
                        ok = aDroid(psDroid).order.typ <> DORDER_PATROL And aDroid(psDroid).order.typ <> DORDER_CIRCLE _
                             And aDroid(psDroid).action <> DACTION_MOVETOREARM And aDroid(psDroid).action <> DACTION_MOVETOREARMPOINT
                        If ok Then ok = actionVTOLLandingPos(psDroid, tar)  ' Can find a sensible place to land.
                        If ok Then ok = map_coord(tar.x) <> map_coord(aDroid(psDroid).sMove.destination.x) Or map_coord(tar.y) <> map_coord(aDroid(psDroid).sMove.destination.y)  ' We're not at the right place to land.
                        If ok Then
                            aDroid(psDroid).sMove.destination = tar
                            moveDroidTo psDroid, aDroid(psDroid).sMove.destination.x, aDroid(psDroid).sMove.destination.y
                        Else
                            aDroid(psDroid).sMove.Status = MOVEHOVER
                        End If
                    Else
                        aDroid(psDroid).sMove.Status = MOVETURN
                    End If
                    objTrace aDroid(psDroid).id, "moveUpdateDroid", "Arrived at destination!"
                    GoTo switchEnd
                End If
            End If

            If aDroid(psDroid).sMove.psFormation <> 0 And (aDroid(psDroid).sMove.pathIndex = aDroid(psDroid).sMove.asPath_size) Then
                ' TODO: What this *used* to do was essentially:
'               SDWORD  fx, fy;
'               if (formationGetPos(psDroid->sMove.psFormation, psDroid, &fx,&fy,true))
'               {
'                   psDroid->sMove.target = Vector2i(fx, fy);
'                   moveCalcBoundary(psDroid);
'                   // See: https://github.com/Warzone2100/warzone2100/commit/4bbd0e9b5fb972d4049456a9c89c3caf2829c4a0#diff-a3ba5fb329c14069f6d815e17db0ac6144cc42aca71effe59d026ddc9278d800
'                   // which removed boundary vectors (which would have picked up this target change via moveCalcBoundary?)
'               }
            End If

            moveDir = moveGetDirection(psDroid)
            moveSpeed = moveCalcDroidSpeed(psDroid)

            If (aDroid(psDroid).sMove.bumpTime <> 0) And _
               ULT(aDroid(psDroid).sMove.pauseTime + aDroid(psDroid).sMove.bumpTime + BLOCK_PAUSETIME, gameTime) Then
                If aDroid(psDroid).sMove.Status = MOVEPOINTTOPOINT Then
                    aDroid(psDroid).sMove.Status = MOVEPAUSE
                Else
                    aDroid(psDroid).sMove.Status = MOVEPOINTTOPOINT
                End If
                aDroid(psDroid).sMove.pauseTime = (gameTime - aDroid(psDroid).sMove.bumpTime) And &HFFFF&
            End If

            If (aDroid(psDroid).sMove.Status = MOVEPAUSE) And _
               (aDroid(psDroid).sMove.bumpTime <> 0) And _
               (aDroid(psDroid).sMove.lastBump > aDroid(psDroid).sMove.pauseTime) And _
               ULT(aDroid(psDroid).sMove.lastBump + aDroid(psDroid).sMove.bumpTime + BLOCK_PAUSERELEASE, gameTime) Then
                aDroid(psDroid).sMove.Status = MOVEPOINTTOPOINT
            End If
        End If
    ElseIf status = MOVETURN Then
        ' Turn the droid to it's final facing
        turning = False
        If aDroid(psDroid).sMove.psFormation <> 0 Then
            turning = aFormation(aDroid(psDroid).sMove.psFormation).refCount > 1 _
                      And aDroid(psDroid).rot.direction <> aFormation(aDroid(psDroid).sMove.psFormation).direction
        End If
        If turning Then
            moveSpeed = 0
            moveDir = aFormation(aDroid(psDroid).sMove.psFormation).direction
        Else
            If asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_LIFT Then
                aDroid(psDroid).sMove.Status = MOVEPOINTTOPOINT
            Else
                aDroid(psDroid).sMove.Status = MOVEINACTIVE
            End If
        End If
    ElseIf status = MOVETURNTOTARGET Then
        moveSpeed = 0
        moveDir = iAtan2(aDroid(psDroid).sMove.target.x - aDroid(psDroid).pos.x, aDroid(psDroid).sMove.target.y - aDroid(psDroid).pos.y)
    ElseIf status = MOVEHOVER Then
        moveDescending psDroid
    Else
        ASSERT_FAILURE "false", "src/move.cpp:2384", "moveUpdateDroid", "unknown move state"
        Exit Sub
    End If
switchEnd:

    ' Update the movement model for the droid
    oldx = aDroid(psDroid).pos.x
    oldy = aDroid(psDroid).pos.y

    If aDroid(psDroid).droidType = DROID_PERSON Then
        moveUpdatePersonModel psDroid, moveSpeed, moveDir
    ElseIf DROID_isCyborg(psDroid) Then
        moveUpdateCyborgModel psDroid, moveSpeed, moveDir, oldStatus
    ElseIf asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_LIFT Then
        moveUpdateVtolModel psDroid, moveSpeed, moveDir
    Else
        moveUpdateGroundModel psDroid, moveSpeed, moveDir
    End If

    If map_coord(oldx) <> map_coord(aDroid(psDroid).pos.x) _
       Or map_coord(oldy) <> map_coord(aDroid(psDroid).pos.y) Then
        visTilesUpdate psDroid

        ' object moved from one tile to next, check to see if droid is near stuff.(oil)
        checkLocalFeatures psDroid

        triggerEventDroidMoved psDroid, oldx, oldy
    End If

    ' See if it's got blocked
    If (asPropulsionStats(psPropStats).propulsionType <> PROPULSION_TYPE_LIFT) Then
        If moveBlocked(psDroid) Then
            objTrace aDroid(psDroid).id, "moveUpdateDroid", "status: id %d blocked", aDroid(psDroid).id
            aDroid(psDroid).sMove.Status = MOVETURN
        End If
    End If

'   // If were in drive mode and the droid is a follower then stop it when it gets within
'   // range of the driver.
'   if(driveIsFollower(psDroid)) {
'       if(DoFollowRangeCheck) {
'           if(driveInDriverRange(psDroid)) {
'               psDroid->sMove.Status = MOVEINACTIVE;
'//             ClearFollowRangeCheck = true;
'           } else {
'               AllInRange = false;
'           }
'       }
'   }

    ' If it's sitting in water then it's got to go with the flow!
    If worldOnMap(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y) Then
        If terrainType(mapTile(map_coord(aDroid(psDroid).pos.x), map_coord(aDroid(psDroid).pos.y))) = TER_WATER Then
            updateDroidOrientation psDroid
        End If
    End If

    If aDroid(psDroid).sMove.Status = MOVETURNTOTARGET And aDroid(psDroid).rot.direction = moveDir Then
        If asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_LIFT Then
            aDroid(psDroid).sMove.Status = MOVEPOINTTOPOINT
        Else
            aDroid(psDroid).sMove.Status = MOVEINACTIVE
        End If
        objTrace aDroid(psDroid).id, "moveUpdateDroid", "MOVETURNTOTARGET complete"
    End If

    If aDroid(psDroid).periodicalDamageStart <> 0 And aDroid(psDroid).droidType <> DROID_PERSON And BASE_OBJECT_visibleForLocalDisplay(psDroid) <> 0 Then ' display-only check for adding effect
        pos.x = aDroid(psDroid).pos.x + (18 - rand() Mod 36)
        pos.z = aDroid(psDroid).pos.y + (18 - rand() Mod 36)
        pos.y = aDroid(psDroid).pos.z + (imdMaxY(aDroid(psDroid).sDisplay.imd) \ 3)
        addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_SMALL, False, 0, 0, gameTime - deltaGameTime + 1
    End If

    movePlayAudio psDroid, bStarted, bStopped, moveSpeed
    If Not (droidOnMap(psDroid)) Then ASSERT_FAILURE "droidOnMap(psDroid)", "src/move.cpp:2469", "moveUpdateDroid", "%s moved off map (%u, %u)->(%u, %u)", droidGetName(psDroid), oldx, oldy, aDroid(psDroid).pos.x, aDroid(psDroid).pos.y
    checkDroid psDroid, "src/move.cpp:2470", "moveUpdateDroid", max_check_object_recursion
End Sub
