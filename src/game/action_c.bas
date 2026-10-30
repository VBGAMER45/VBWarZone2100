Attribute VB_Name = "action_c"
' src/action.cpp, action.h: Functions for setting the action of a droid.
' (VB) A tileMatchFunction is a number (TILE_MATCH_...) and a Select Case;
' its matchState is this module's variables. actionUpdateDroid's
' actionUpdateFunc is a number (UPDATE_FUNC_...). actionTargetTurret takes
' the weapon's slot where C++ takes WEAPON *psWeapon (which is always
' &psAttacker->asWeaps[slot]; C++ gets the slot back by subtracting).
' The overloads of actionDroid: actionDroid(psDroid, action),
' actionDroid_loc(.., x, y), actionDroid_obj(.., psObj),
' actionDroid_objLoc(.., psObj, x, y).
Option Explicit

' attack run distance
Private Const VTOL_ATTACK_LENGTH As Long = 1000
Private Const VTOL_ATTACK_TARDIST As Long = 400

' turret rotation limit (DEG(45), DEG(60))
Private Const VTOL_TURRET_LIMIT As Long = 8192
Private Const VTOL_TURRET_LIMIT_BOMB As Long = 10922

Private Const VTOL_ATTACK_AUDIO_DELAY As Long = 3 * GAME_TICKS_PER_SEC

' Droids heavier than this rotate and pitch more slowly.
Private Const HEAVY_WEAPON_WEIGHT As Long = 50000

Private Const ACTION_TURRET_ROTATION_RATE As Long = 45
Private Const REPAIR_PITCH_LOWER As Long = 30
Private Const REPAIR_PITCH_UPPER As Long = -15

' How many tiles to pull back.
Private Const PULL_BACK_DIST As Long = 10

' action.h
' after failing a route ... this is the amount of time that the droid goes all defensive until it can start going aggressive
Public Const MIN_SULK_TIME As Long = 1500       ' 1.5 sec
Public Const MAX_SULK_TIME As Long = 4000       ' 4 secs

' this is how far a droid will move to sit next to a repair facility to be repaired
Public Const REPAIR_RANGE As Long = 2 * TILE_UNITS

' this is how far a droid will move to sit next to another droid it wants to repair
Public Const REPAIR_MAXDIST As Long = 5 * TILE_UNITS

' minimum structure strength modifier needed to automatically target blocking walls
Public Const MIN_STRUCTURE_BLOCK_STRENGTH As Long = 33

' How many ticks to skip before looking for a better target.
Public Const TARGET_UPD_SKIP_FRAMES As Long = 1000
Public Const TARGET_CHECK_NEW_SKIP_TICKS As Long = 200

Public Const EMP_DISABLE_TIME As Long = 10000     ' 10 secs

' weapondef.h: DEG(1)
Public Const FIXED_TURRET_DIR As Long = 182

' data required for any action
Private Type DROID_ACTION_DATA
    action As DROID_ACTION
    x As Long                           ' UDWORD
    y As Long
    ' multiple action target info
    psObj As Long                       ' BASE_OBJECT *
    psStats As Long                     ' BASE_STATS *
End Type

' Radius for search when looking for VTOL landing position
Private Const vtolLandingRadius As Long = 23

' (VB) the tileMatchFunctions
Private Const TILE_MATCH_VTOL_LANDING As Long = 1  ' vtolLandingTileSearchFunction

' (VB) actionUpdateDroid's actionUpdateFunc
Private Const UPDATE_FUNC_NONE As Long = 0
Private Const UPDATE_FUNC_DEMOLISHING As Long = 1  ' droidUpdateDemolishing
Private Const UPDATE_FUNC_REPAIR As Long = 2       ' droidUpdateRepair
Private Const UPDATE_FUNC_RESTORE As Long = 3      ' droidUpdateRestore

Private landingXY As Vector2i           ' vtolLandingTileSearchFunction's matchState (Vector2i *)

Private gridListBuildPos() As Long      ' actionRemoveDroidsFromBuildPos's static GridList

' Check if a droid has stopped moving
Public Function DROID_STOPPED(ByVal psDroid As Long) As Boolean
    DROID_STOPPED = (aDroid(psDroid).sMove.Status = MOVEINACTIVE Or aDroid(psDroid).sMove.Status = MOVEHOVER Or _
                     aDroid(psDroid).sMove.Status = MOVESHUFFLE)
End Function

' Throttle how often we look for a new target
' Notes: deltaGameTime is either 0 (if no gameTime update was processed) or GAME_TICKS_PER_UPDATE
Public Function IS_TIME_TO_CHECK_FOR_NEW_TARGET(ByVal psDroid As Long) As Boolean
    IS_TIME_TO_CHECK_FOR_NEW_TARGET = Fix(U32(FromU32(CDbl(aDroid(psDroid).id) + gameTime)) / TARGET_CHECK_NEW_SKIP_TICKS) <> _
                                      Fix(U32(FromU32(CDbl(aDroid(psDroid).id) + gameTime - deltaGameTime)) / TARGET_CHECK_NEW_SKIP_TICKS)
End Function

Public Function getDroidActionName(ByVal action As DROID_ACTION) As String
    Static name_() As String
    If (Not Not name_) = 0 Then
        name_ = Split("DACTION_NONE|DACTION_MOVE|DACTION_BUILD|DACTION_BUILD_FOUNDATION|DACTION_DEMOLISH|DACTION_REPAIR|" & _
                     "DACTION_ATTACK|DACTION_OBSERVE|DACTION_FIRESUPPORT|DACTION_SULK|DACTION_DESTRUCT|DACTION_TRANSPORTOUT|" & _
                     "DACTION_TRANSPORTWAITTOFLYIN|DACTION_TRANSPORTIN|DACTION_DROIDREPAIR|DACTION_RESTORE|DACTION_UNUSED|" & _
                     "DACTION_MOVEFIRE|DACTION_MOVETOBUILD|DACTION_MOVETODEMOLISH|DACTION_MOVETOREPAIR|DACTION_BUILDWANDER|" & _
                     "DACTION_FOUNDATION_WANDER|DACTION_MOVETOATTACK|DACTION_ROTATETOATTACK|DACTION_MOVETOOBSERVE|" & _
                     "DACTION_WAITFORREPAIR|DACTION_MOVETOREPAIRPOINT|DACTION_WAITDURINGREPAIR|DACTION_MOVETODROIDREPAIR|" & _
                     "DACTION_MOVETORESTORE|DACTION_UNUSED2|DACTION_MOVETOREARM|DACTION_WAITFORREARM|DACTION_MOVETOREARMPOINT|" & _
                     "DACTION_WAITDURINGREARM|DACTION_VTOLATTACK|DACTION_CLEARREARMPAD|DACTION_RETURNTOPOS|" & _
                     "DACTION_FIRESUPPORT_RETREAT|ACTION UNKNOWN|DACTION_CIRCLE", "|")
    End If

    If Not (ULT(action, UBound(name_) + 1)) Then ASSERT_FAILURE "action < sizeof(name) / sizeof(name[0])", "src/action.cpp:155", "getDroidActionName", "DROID_ACTION out of range: %u", action: getDroidActionName = vbNullString: Exit Function

    getDroidActionName = name_(action)
End Function

' When useLongWithOptimum is true: Returns true if distance from psDroid to psObj is within optimum range
' When useLongWithOptimum is false: Returns true if if distance from psDroid to psObj is within long range
Public Function actionInRange(ByVal psDroid As Long, ByVal psObj As Long, ByVal weapon_slot As Long, Optional ByVal useLongWithOptimum As Boolean = True) As Boolean
    Dim compIndex As Long, dx As Long, dy As Long, radSq As Long, longRange As Long, shortRange As Long, rangeSq As Long, minrange As Long
    Dim objPos As Vector3i

    checkDroid psDroid, "src/action.cpp:166", "actionInRange", max_check_object_recursion

    If aDroid(psDroid).asWeaps(0).nStat = 0 Then
        actionInRange = False
        Exit Function
    End If

    compIndex = aDroid(psDroid).asWeaps(weapon_slot).nStat
    If Not (ULT(compIndex, asWeaponStats_size)) Then ASSERT_FAILURE "compIndex < asWeaponStats.size()", "src/action.cpp:174", "actionInRange", "Invalid range referenced for numWeaponStats, %d > %zu", compIndex, asWeaponStats_size: actionInRange = False: Exit Function

    objPos = SIMPLE_OBJECT_pos(psObj)
    dx = aDroid(psDroid).pos.x - objPos.x
    dy = aDroid(psDroid).pos.y - objPos.y

    radSq = dx * dx + dy * dy
    longRange = proj_GetLongRange(asWeaponStats(compIndex), aDroid(psDroid).player)
    shortRange = proj_GetShortRange(asWeaponStats(compIndex), aDroid(psDroid).player)

    rangeSq = 0
    Select Case aDroid(psDroid).secondaryOrder And DSS_ARANGE_MASK
    Case DSS_ARANGE_OPTIMUM
        If Not useLongWithOptimum And weaponShortHit(asWeaponStats(compIndex), aDroid(psDroid).player) > weaponLongHit(asWeaponStats(compIndex), aDroid(psDroid).player) Then
            rangeSq = shortRange * shortRange
        Else
            rangeSq = longRange * longRange
        End If
    Case DSS_ARANGE_SHORT
        rangeSq = shortRange * shortRange
    Case DSS_ARANGE_LONG
        rangeSq = longRange * longRange
    Case Else
        ASSERT_FAILURE "!""unknown attackrange order""", "src/action.cpp:204", "actionInRange", "unknown attack range order"
        rangeSq = longRange * longRange
    End Select

    ' check max range
    If radSq <= rangeSq Then
        ' check min range
        minrange = proj_GetMinRange(asWeaponStats(compIndex), aDroid(psDroid).player)
        If radSq >= minrange * minrange Then
            actionInRange = True
            Exit Function
        ElseIf Not proj_Direct(compIndex) Then
            actionInRange = True
            Exit Function
        End If
    End If

    actionInRange = False
End Function

' Retuns true if distance from psDroid to target is less than, or equal to droid's
' minimum range
Private Function actionInsideMinRange(ByVal psDroid As Long, ByVal psObj As Long, ByVal psStats As Long) As Boolean
    Dim dx As Long, dy As Long, radSq As Long, minRange As Long, rangeSq As Long, objPos As Vector3i

    checkDroid psDroid, "src/action.cpp:229", "actionInsideMinRange", max_check_object_recursion
    checkObject psObj, "src/action.cpp:230", "actionInsideMinRange", max_check_object_recursion

    If psStats = -1 Then
        psStats = DROID_getWeaponStats(psDroid, 0)
    End If

    ' if I am a multi-turret droid
    If aDroid(psDroid).asWeaps(0).nStat = 0 Then
        actionInsideMinRange = False
        Exit Function
    End If

    objPos = SIMPLE_OBJECT_pos(psObj)
    dx = aDroid(psDroid).pos.x - objPos.x
    dy = aDroid(psDroid).pos.y - objPos.y
    radSq = dx * dx + dy * dy
    minRange = proj_GetMinRange(asWeaponStats(psStats), aDroid(psDroid).player)
    rangeSq = minRange * minRange

    ' check min range
    If radSq <= rangeSq Then
        actionInsideMinRange = True
        Exit Function
    End If

    actionInsideMinRange = False
End Function

' Realign turret
Public Sub actionAlignTurret(ByVal psObj As Long, ByVal weapon_slot As Long)
    Dim nearest As Long, tRot As Long, tPitch As Long, rotationFrame As Long

    nearest = 0

    'get the maximum rotation this frame
    rotationFrame = gameTimeAdjustedIncrement(DEG(ACTION_TURRET_ROTATION_RATE))

    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        tRot = aDroid(psObj).asWeaps(weapon_slot).rot.direction
        tPitch = aDroid(psObj).asWeaps(weapon_slot).rot.pitch
    Case OBJ_STRUCTURE
        tRot = aStruct(psObj).asWeaps(weapon_slot).rot.direction
        tPitch = aStruct(psObj).asWeaps(weapon_slot).rot.pitch

        ' now find the nearest 90 degree angle
        nearest = (((tRot + DEG(45)) \ DEG(90)) * DEG(90)) And &HFFFF&  ' Cast wrapping intended.
    Case Else
        ASSERT_FAILURE "!""invalid object type""", "src/action.cpp:283", "actionAlignTurret", "invalid object type"
        Exit Sub
    End Select

    tRot = (tRot + clip_(angleDelta(nearest - tRot), -rotationFrame, rotationFrame)) And &HFFFF&  ' Addition wrapping intended.

    ' align the turret pitch
    tPitch = (tPitch + clip_(angleDelta(0 - tPitch), -rotationFrame \ 2, rotationFrame \ 2)) And &HFFFF&  ' Addition wrapping intended.

    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        aDroid(psObj).asWeaps(weapon_slot).rot.direction = tRot
        aDroid(psObj).asWeaps(weapon_slot).rot.pitch = tPitch
    Case OBJ_STRUCTURE
        aStruct(psObj).asWeaps(weapon_slot).rot.direction = tRot
        aStruct(psObj).asWeaps(weapon_slot).rot.pitch = tPitch
    Case Else
        ASSERT_FAILURE "!""invalid object type""", "src/action.cpp:303", "actionAlignTurret", "invalid object type"
        Exit Sub
    End Select
End Sub

' returns true if on target
' (VB) weapon_slot: psWeapon is &psAttacker->asWeaps[weapon_slot]
Public Function actionTargetTurret(ByVal psAttacker As Long, ByVal psTarget As Long, ByVal weapon_slot As Long) As Boolean
    Dim psWeapon As WEAPON, psWeapStats As Long, tRotation As Long, tPitch As Long, targetRotation As Long
    Dim rotationTolerance As Long, pitchLowerLimit As Long, pitchUpperLimit As Long, bRepair As Boolean
    Dim rotRate As Long, pitchRate As Long, excess As Long, attackerTurretPos As Vector3i, attackerPos As Vector3i, attackerRot As Rotation
    Dim targetPos As Vector3i, rotationError As Long, limit As Long, onTarget As Boolean, minRange As Long
    Dim delta As Vector3i, dxy As Long, targetPitch As Long, pitchError As Long

    If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
        psWeapon = aDroid(psAttacker).asWeaps(weapon_slot)
    Else
        psWeapon = aStruct(psAttacker).asWeaps(weapon_slot)
    End If
    psWeapStats = psWeapon.nStat
    rotationTolerance = 0

    If psTarget = 0 Then
        actionTargetTurret = False
        Exit Function
    End If

    If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
        bRepair = aDroid(psAttacker).droidType = DROID_REPAIR
    End If

    ' these are constants now and can be set up at the start of the function
    rotRate = DEG(ACTION_TURRET_ROTATION_RATE) * 4
    pitchRate = DEG(ACTION_TURRET_ROTATION_RATE) * 2

    ' extra heavy weapons on some structures need to rotate and pitch more slowly
    If asWeaponStats(psWeapStats).weight > HEAVY_WEAPON_WEIGHT And Not bRepair Then
        excess = FromU32(Fix(U32(MUL32(DEG(100), asWeaponStats(psWeapStats).weight - HEAVY_WEAPON_WEIGHT)) / U32(asWeaponStats(psWeapStats).weight)))

        rotRate = FromU32(U32(DEG(ACTION_TURRET_ROTATION_RATE) * 2) - U32(excess))
        pitchRate = rotRate \ 2
    End If

    tRotation = psWeapon.rot.direction
    tPitch = psWeapon.rot.pitch

    'set the pitch limits based on the weapon stats of the attacker
    pitchLowerLimit = 0
    pitchUpperLimit = 0
    attackerPos = SIMPLE_OBJECT_pos(psAttacker)
    attackerRot = SIMPLE_OBJECT_rot(psAttacker)
    attackerTurretPos = attackerPos  ' Using for calculating the pitch, but not the direction, in case using the exact direction causes bugs somewhere.
    If SIMPLE_OBJECT_type(psAttacker) = OBJ_STRUCTURE Then
        calcStructureMuzzleBaseLocation psAttacker, attackerTurretPos, weapon_slot
        pitchLowerLimit = DEG(asWeaponStats(psWeapStats).minElevation)
        pitchUpperLimit = DEG(asWeaponStats(psWeapStats).maxElevation)
    ElseIf SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
        calcDroidMuzzleBaseLocation psAttacker, attackerTurretPos, weapon_slot

        If aDroid(psAttacker).droidType = DROID_WEAPON Or DROID_isTransporter(psAttacker) _
           Or aDroid(psAttacker).droidType = DROID_COMMAND Or aDroid(psAttacker).droidType = DROID_CYBORG _
           Or aDroid(psAttacker).droidType = DROID_CYBORG_SUPER Then
            pitchLowerLimit = DEG(asWeaponStats(psWeapStats).minElevation)
            pitchUpperLimit = DEG(asWeaponStats(psWeapStats).maxElevation)
        ElseIf aDroid(psAttacker).droidType = DROID_REPAIR Then
            pitchLowerLimit = DEG(REPAIR_PITCH_LOWER)
            pitchUpperLimit = DEG(REPAIR_PITCH_UPPER)
        End If
    End If

    'get the maximum rotation this frame
    rotRate = gameTimeAdjustedIncrement(rotRate)
    rotRate = MAX_(rotRate, DEG(1))
    pitchRate = gameTimeAdjustedIncrement(pitchRate)
    pitchRate = MAX_(pitchRate, DEG(1))

    'and point the turret at target
    targetPos = SIMPLE_OBJECT_pos(psTarget)
    targetRotation = calcDirection(attackerPos.x, attackerPos.y, targetPos.x, targetPos.y)

    'restrict rotationerror to =/- 180 degrees
    rotationError = angleDelta(targetRotation - (tRotation + attackerRot.direction))

    tRotation = (tRotation + clip_(rotationError, -rotRate, rotRate)) And &HFFFF&  ' Addition wrapping intentional.
    If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
        If DROID_isVtol(psAttacker) Then
            ' limit the rotation for vtols
            limit = VTOL_TURRET_LIMIT
            If asWeaponStats(psWeapStats).weaponSubClass = WSC_BOMB Or asWeaponStats(psWeapStats).weaponSubClass = WSC_EMP Then
                limit = 0  ' Don't turn bombs.
                rotationTolerance = VTOL_TURRET_LIMIT_BOMB
            End If
            tRotation = clip_(angleDelta(tRotation), -limit, limit) And &HFFFF&  ' Cast wrapping intentional.
        End If
    End If
    onTarget = Abs(angleDelta(targetRotation - (tRotation + attackerRot.direction))) <= rotationTolerance

    ' Set muzzle pitch if not repairing or outside minimum range
    minRange = proj_GetMinRange(asWeaponStats(psWeapStats), SIMPLE_OBJECT_player(psAttacker))
    If Not bRepair Then
        If ULT(MUL32(minRange, minRange), objPosDiffSq(attackerPos, targetPos)) Then
            If proj_Direct(psWeapStats) Then
                ' get target distance
                delta.x = targetPos.x - attackerTurretPos.x
                delta.y = targetPos.y - attackerTurretPos.y
                delta.z = targetPos.z - attackerTurretPos.z
                dxy = iHypot(delta.x, delta.y)

                targetPitch = iAtan2(delta.z, dxy) And &HFFFF&
                targetPitch = clip_(angleDelta(targetPitch), pitchLowerLimit, pitchUpperLimit) And &HFFFF&  ' Cast wrapping intended.
                pitchError = angleDelta(targetPitch - tPitch)

                tPitch = (tPitch + clip_(pitchError, -pitchRate, pitchRate)) And &HFFFF&  ' Addition wrapping intended.
                onTarget = onTarget And targetPitch = tPitch
            End If
        End If
    End If

    If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
        aDroid(psAttacker).asWeaps(weapon_slot).rot.direction = tRotation
        aDroid(psAttacker).asWeaps(weapon_slot).rot.pitch = tPitch
    Else
        aStruct(psAttacker).asWeaps(weapon_slot).rot.direction = tRotation
        aStruct(psAttacker).asWeaps(weapon_slot).rot.pitch = tPitch
    End If

    actionTargetTurret = onTarget
End Function

' return whether a droid can see a target to fire on it
Public Function actionVisibleTarget(ByVal psDroid As Long, ByVal psTarget As Long, ByVal weapon_slot As Long) As Boolean
    checkDroid psDroid, "src/action.cpp:423", "actionVisibleTarget", max_check_object_recursion
    If Not (psTarget <> 0) Then ASSERT_FAILURE "psTarget != nullptr", "src/action.cpp:424", "actionVisibleTarget", "Target is NULL": actionVisibleTarget = False: Exit Function
    If Not (aDroid(psDroid).player < MAX_PLAYERS) Then ASSERT_FAILURE "psDroid->player < MAX_PLAYERS", "src/action.cpp:425", "actionVisibleTarget", "psDroid->player (%u) must be < MAX_PLAYERS", aDroid(psDroid).player: actionVisibleTarget = False: Exit Function
    If BASE_OBJECT_visible(psTarget, aDroid(psDroid).player) = 0 Then
        actionVisibleTarget = False
        Exit Function
    End If
    If aDroid(psDroid).numWeaps = 0 Or DROID_isVtol(psDroid) Then
        If visibleObject(psDroid, psTarget, False) <> 0 Then
            actionVisibleTarget = True
            Exit Function
        End If
    End If
    If orderState(psDroid, DORDER_FIRESUPPORT) Then
        actionVisibleTarget = lineOfFire(psDroid, psTarget, weapon_slot, True)
    ElseIf visibleObject(psDroid, psTarget, False) > UBYTE_MAX \ 2 Then
        actionVisibleTarget = lineOfFire(psDroid, psTarget, weapon_slot, True)
    Else
        actionVisibleTarget = False
    End If
End Function

Private Sub actionAddVtolAttackRun(ByVal psDroid As Long)
    Dim psTarget As Long, delta As Vector2i, dist As Long, dest As Vector2i, targetPos As Vector3i

    checkDroid psDroid, "src/action.cpp:442", "actionAddVtolAttackRun", max_check_object_recursion

    If aDroid(psDroid).psActionTarget(0) <> 0 Then
        psTarget = aDroid(psDroid).psActionTarget(0)
    ElseIf aDroid(psDroid).order.psObj <> 0 Then
        psTarget = aDroid(psDroid).order.psObj
    Else
        Exit Sub
    End If

    ' get normal vector from droid to target
    targetPos = SIMPLE_OBJECT_pos(psTarget)
    delta.x = targetPos.x - aDroid(psDroid).pos.x
    delta.y = targetPos.y - aDroid(psDroid).pos.y

    ' get magnitude of normal vector (Pythagorean theorem)
    dist = MAX_(iHypot(delta.x, delta.y), 1)

    ' add waypoint behind target attack length away
    dest.x = targetPos.x + delta.x * VTOL_ATTACK_LENGTH \ dist
    dest.y = targetPos.y + delta.y * VTOL_ATTACK_LENGTH \ dist

    If Not worldOnMap(dest.x, dest.y) Then
        debug_ 468, LOG_NEVER, "actionAddVtolAttackRun", "*** actionAddVtolAttackRun: run off map! ***"
    Else
        moveDroidToDirect psDroid, dest.x, dest.y
    End If
End Sub

Private Sub actionUpdateVtolAttack(ByVal psDroid As Long)
    checkDroid psDroid, "src/action.cpp:478", "actionUpdateVtolAttack", max_check_object_recursion

    ' don't do attack runs whilst returning to base
    If aDroid(psDroid).order.typ = DORDER_RTB Then
        Exit Sub
    End If

    ' order back to base after fixed number of attack runs
    If aDroid(psDroid).numWeaps > 0 And aDroid(psDroid).asWeaps(0).nStat > 0 Then
        If vtolEmpty(psDroid) Then
            moveToRearm psDroid
            Exit Sub
        End If
    End If

    ' circle around target if hovering and not cyborg
    If aDroid(psDroid).sMove.Status = MOVEHOVER And Not DROID_isCyborg(psDroid) Then
        actionAddVtolAttackRun psDroid
    End If
End Sub

Private Sub actionUpdateTransporter(ByVal psDroid As Long)
    checkDroid psDroid, "src/action.cpp:502", "actionUpdateTransporter", max_check_object_recursion

    'check if transporter has arrived
    If updateTransporter(psDroid) Then
        ' Got to destination
        aDroid(psDroid).action = DACTION_NONE
    End If
End Sub

' calculate a position for units to pull back to if they
' need to increase the range between them and a target
Private Sub actionCalcPullBackPoint(ByVal psObj As Long, ByVal psTarget As Long, px As Long, py As Long)
    Dim xdiff As Long, ydiff As Long, len_ As Long, objPos As Vector3i, targetPos As Vector3i

    ' get the vector from the target to the object
    objPos = SIMPLE_OBJECT_pos(psObj)
    targetPos = SIMPLE_OBJECT_pos(psTarget)
    xdiff = objPos.x - targetPos.x
    ydiff = objPos.y - targetPos.y
    len_ = iHypot(xdiff, ydiff)

    If len_ = 0 Then
        xdiff = TILE_UNITS
        ydiff = TILE_UNITS
    Else
        xdiff = (xdiff * TILE_UNITS) \ len_
        ydiff = (ydiff * TILE_UNITS) \ len_
    End If

    ' create the position
    px = objPos.x + xdiff * PULL_BACK_DIST
    py = objPos.y + ydiff * PULL_BACK_DIST

    ' make sure coordinates stay inside of the map
    clip_world_offmap px, py
End Sub

' check whether a droid is in the neighboring tile of another droid
Public Function actionReachedDroid(ByVal psDroid As Long, ByVal psOther As Long) As Boolean
    Dim delta As Vector2i
    If Not (psDroid <> 0 And psOther <> 0) Then ASSERT_FAILURE "psDroid != nullptr && psOther != nullptr", "src/action.cpp:544", "actionReachedDroid", "Bad droids": actionReachedDroid = False: Exit Function
    checkDroid psDroid, "src/action.cpp:545", "actionReachedDroid", max_check_object_recursion
    delta.x = map_coord(aDroid(psDroid).pos.x) - map_coord(aDroid(psOther).pos.x)
    delta.y = map_coord(aDroid(psDroid).pos.y) - map_coord(aDroid(psOther).pos.y)
    actionReachedDroid = delta.x >= -1 And delta.x <= 1 And delta.y >= -1 And delta.y <= 1
End Function

' check whether a droid is in the neighboring tile to a build position
' psStats: a BASE_STATS * handle
Public Function actionReachedBuildPos(ByVal psDroid As Long, ByVal x As Long, ByVal y As Long, ByVal dir_ As Long, ByVal psStats As Long) As Boolean
    Dim b As StructureBounds, delta As Vector2i

    If Not (psStats <> 0 And psDroid <> 0) Then ASSERT_FAILURE "psStats != nullptr && psDroid != nullptr", "src/action.cpp:556", "actionReachedBuildPos", "Bad stat or droid": actionReachedBuildPos = False: Exit Function
    checkDroid psDroid, "src/action.cpp:557", "actionReachedBuildPos", max_check_object_recursion

    b = getStructureBounds_stats(psStats, Vector2i_(x, y), dir_)

    ' do all calculations in half tile units so that
    ' the droid moves to within half a tile of the target
    ' NOT ANY MORE - JOHN
    delta.x = map_coord(aDroid(psDroid).pos.x) - b.map.x
    delta.y = map_coord(aDroid(psDroid).pos.y) - b.map.y
    actionReachedBuildPos = delta.x >= -1 And delta.x <= b.size.x And delta.y >= -1 And delta.y <= b.size.y
End Function

' check if a droid is on the foundations of a new building
' psStats: a BASE_STATS * handle
Private Function actionRemoveDroidsFromBuildPos(ByVal player As Long, pos As Vector2i, ByVal dir_ As Long, ByVal psStats As Long) As Boolean
    Dim buildPosEmpty As Boolean, b As StructureBounds, structureCentre As Vector2i, structureMaxRadius As Long
    Dim n As Long, gi As Long, droid_ As Long, delta As Vector2i, bestDest As Vector2i, bestDist As Double
    Dim x As Long, y As Long, dest As Vector2i, dist As Long, newPos As Vector2i, ang As Long

    If Not (psStats <> 0) Then ASSERT_FAILURE "psStats != nullptr", "src/action.cpp:572", "actionRemoveDroidsFromBuildPos", "Bad stat": actionRemoveDroidsFromBuildPos = False: Exit Function

    buildPosEmpty = True

    b = getStructureBounds_stats(psStats, pos, dir_)

    structureCentre.x = world_coord(b.map.x) + world_coord(b.size.x) \ 2
    structureCentre.y = world_coord(b.map.y) + world_coord(b.size.y) \ 2
    structureMaxRadius = iHypot(world_coord(b.size.x) \ 2, world_coord(b.size.y) \ 2) + 1  ' +1 since iHypot rounds down.

    n = gridStartIterate(structureCentre.x, structureCentre.y, structureMaxRadius, gridListBuildPos)
    For gi = 0 To n - 1
        droid_ = castDroid(gridListBuildPos(gi))
        If droid_ = 0 Then
            GoTo continue_  ' Only looking for droids.
        End If

        delta.x = map_coord(aDroid(droid_).pos.x) - b.map.x
        delta.y = map_coord(aDroid(droid_).pos.y) - b.map.y
        If delta.x < 0 Or delta.x >= b.size.x Or delta.y < 0 Or delta.y >= b.size.y Then
            GoTo continue_  ' Droid not under new structure (just near it).
        ElseIf DROID_isFlying(droid_) Then
            GoTo continue_
        ElseIf DROID_isFlightBasedTransporter(droid_) Then
            GoTo continue_
        End If

        buildPosEmpty = False  ' Found a droid, have to move it away.

        If Not aiCheckAlliances(player, aDroid(droid_).player) Then
            GoTo continue_  ' Enemy droids probably don't feel like moving.
        End If

        ' TODO If the action code was less convoluted, it would be possible for the droid should drive away instead of just getting moved away.
        bestDest = Vector2i_(0, 0)  ' Dummy initialisation.
        bestDist = 4294967295#      ' UINT32_MAX
        For y = -1 To b.size.y
            x = -1
            Do While x <= b.size.x
                dest.x = world_coord(b.map.x + x) + TILE_UNITS \ 2
                dest.y = world_coord(b.map.y + y) + TILE_UNITS \ 2
                dist = iHypot(aDroid(droid_).pos.x - dest.x, aDroid(droid_).pos.y - dest.y)
                If dist < bestDist Then
                    If Not fpathBlockingTile(map_coord(dest.x), map_coord(dest.y), asPropulsionStats(DROID_getPropulsionStats(droid_)).propulsionType) Then
                        bestDest = dest
                        bestDist = dist
                    End If
                End If
                If y >= 0 And y < b.size.y Then x = x + b.size.x + 1 Else x = x + 1
            Loop
        Next
        If bestDist <> 4294967295# Then
            ' Push the droid out of the way.
            ang = iAtan2(bestDest.x - aDroid(droid_).pos.x, bestDest.y - aDroid(droid_).pos.y)
            newPos.x = aDroid(droid_).pos.x + iSinR(ang, gameTimeAdjustedIncrement(TILE_UNITS))
            newPos.y = aDroid(droid_).pos.y + iCosR(ang, gameTimeAdjustedIncrement(TILE_UNITS))
            droidSetPosition droid_, newPos.x, newPos.y
        End If
continue_:
    Next

    actionRemoveDroidsFromBuildPos = buildPosEmpty
End Function

Public Sub actionSanity(ByVal psDroid As Long)
    Dim avoidOverkill As Boolean, bDirect As Boolean, i As Long, a As DROID_ACTION, clearIt As Boolean, backToPatrol As Boolean

    ' Don't waste ammo unless given a direct attack order.
    a = aDroid(psDroid).action
    avoidOverkill = aDroid(psDroid).order.typ <> DORDER_ATTACK And _
                    (a = DACTION_ATTACK Or a = DACTION_MOVEFIRE Or a = DACTION_MOVETOATTACK Or _
                     a = DACTION_ROTATETOATTACK Or a = DACTION_VTOLATTACK)
    bDirect = False

    ' clear the target if it has died
    For i = 0 To MAX_WEAPONS - 1
        bDirect = proj_Direct(DROID_getWeaponStats(psDroid, i))
        clearIt = False
        If aDroid(psDroid).psActionTarget(i) <> 0 Then
            If avoidOverkill Then
                clearIt = aiObjectIsProbablyDoomed(aDroid(psDroid).psActionTarget(i), bDirect)
            Else
                clearIt = SIMPLE_OBJECT_died(aDroid(psDroid).psActionTarget(i)) <> 0
            End If
        End If
        If clearIt Then
            If aDroid(psDroid).action = DACTION_DROIDREPAIR Then
                droidRepairStopped castDroid(aDroid(psDroid).psActionTarget(0)), psDroid
            End If
            syncDebugObject_ "actionSanity", aDroid(psDroid).psActionTarget(i), Asc("-")
            setDroidActionTarget psDroid, 0, i
            If i = 0 Then
                If aDroid(psDroid).action <> DACTION_MOVEFIRE And _
                   aDroid(psDroid).action <> DACTION_TRANSPORTIN And _
                   aDroid(psDroid).action <> DACTION_TRANSPORTOUT Then
                    aDroid(psDroid).action = DACTION_NONE
                    ' if VTOL - return to rearm pad if not patrolling
                    If DROID_isVtol(psDroid) Then
                        ' (&& and ||: vtolEmpty, whose CHECK_DROID may log, only for a patrol or circle)
                        backToPatrol = aDroid(psDroid).order.typ = DORDER_PATROL Or aDroid(psDroid).order.typ = DORDER_CIRCLE
                        If backToPatrol Then
                            backToPatrol = Not vtolEmpty(psDroid)
                            If Not backToPatrol Then backToPatrol = (aDroid(psDroid).secondaryOrder And DSS_ALEV_MASK) = DSS_ALEV_NEVER
                        End If
                        If backToPatrol Then
                            ' Back to the patrol.
                            actionDroid_loc psDroid, DACTION_MOVE, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
                        Else
                            moveToRearm psDroid
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

' (VB) actionUpdateFunc(psDroid)
Private Function actionUpdateFunc_(ByVal actionUpdateFunc As Long, ByVal psDroid As Long) As Boolean
    Select Case actionUpdateFunc
    Case UPDATE_FUNC_DEMOLISHING: actionUpdateFunc_ = droidUpdateDemolishing(psDroid)
    Case UPDATE_FUNC_REPAIR: actionUpdateFunc_ = droidUpdateRepair(psDroid)
    Case UPDATE_FUNC_RESTORE: actionUpdateFunc_ = droidUpdateRestore(psDroid)
    End Select
End Function

' (VB) a wall blocking a direct weapon: whether it is worth shooting at it
Private Function wallIsWorthShooting(ByVal psDroid As Long, ByVal blockingWall As Long, ByVal weapEffect As WEAPON_EFFECT) As Boolean
    If Not aiCheckAlliances(aDroid(psDroid).player, aStruct(blockingWall).player) Then
        wallIsWorthShooting = asStructStrengthModifier(weapEffect, asStructureStats(aStruct(blockingWall).pStructureType).strength) >= MIN_STRUCTURE_BLOCK_STRENGTH
    End If
End Function

' Update the action state for a droid
' (VB) DROID_STOPPED(psDroid) && !actionReachedBuildPos(psDroid, actionPos.x, actionPos.y, dir, psStats), short-circuited
' as in C; direction_ is psTarget's rotation when psTarget is not 0 (read only then)
Private Function stoppedShortOfBuildPos(ByVal psDroid As Long, ByVal psTarget As Long, ByVal direction_ As Long, ByVal psStats As Long) As Boolean
    If DROID_STOPPED(psDroid) Then
        If psTarget <> 0 Then direction_ = aStruct(psTarget).rot.direction
        stoppedShortOfBuildPos = Not actionReachedBuildPos(psDroid, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y, direction_, psStats)
    End If
End Function

' (VB) !DROID_STOPPED(psDroid) && Status != MOVETURNTOTARGET && Status != MOVESHUFFLE &&
' actionReachedBuildPos(psDroid, actionPos.x, actionPos.y, psTarget's rotation, psStats), short-circuited as in C
Private Function movingAtBuildPos(ByVal psDroid As Long, ByVal psTarget As Long, ByVal psStats As Long) As Boolean
    If Not DROID_STOPPED(psDroid) Then
        If aDroid(psDroid).sMove.Status <> MOVETURNTOTARGET And aDroid(psDroid).sMove.Status <> MOVESHUFFLE Then
            movingAtBuildPos = actionReachedBuildPos(psDroid, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y, aStruct(psTarget).rot.direction, psStats)
        End If
    End If
End Function

Public Sub actionUpdateDroid(ByVal psDroid As Long)
    Dim actionUpdateFunc As Long
    Dim nonNullWeapon(0 To MAX_WEAPONS - 1) As Boolean
    Dim psTargets(0 To MAX_WEAPONS - 1) As Long
    Dim hasValidWeapon As Boolean, hasVisibleTarget As Boolean
    Dim targetVisibile(0 To MAX_WEAPONS - 1) As Boolean
    Dim bHasTarget As Boolean, bDirect As Boolean, blockingWall As Long, wallBlocked As Boolean
    Dim psPropStats As Long, secHoldActive As Boolean, i As Long, psTemp As Long, psWeapStats As Long
    Dim psActionTarget As Long, compIndex As Long, weapEffect As WEAPON_EFFECT, dirDiff As Long, targetDir As Long
    Dim psTarget As Long, supportsSensorTower As Boolean, commander As Long, targetIsValid As Boolean
    Dim attackRunDelta As Vector2i, diff As Vector2i, rangeSq As Long, maxRange As Long, pos As Vector3i
    Dim chaseBloke As Boolean, actionIsInRange As Boolean, pbx As Long, pby As Long
    Dim desiredStructure As Long, structureAtBuildPosition As Long, droidCannotBuild As Boolean
    Dim buildPosEmpty As Boolean, helpBuild As Boolean, psStructStats As Long, dir_ As Long, bpos As Vector2i
    Dim psStruct As Long, psTile As Long, feature_ As Long, structureAtPos As Long, cantDoRepairLikeAction As Boolean
    Dim ObjAllied As Boolean, ObjDepartedAlly As Boolean, reached As Boolean, xdiff As Long, ydiff As Long
    Dim actionTargetObj As Long, actionTarget As Long, enoughTimeRemaining As Boolean, droidX As Long, droidY As Long
    Dim notify As Boolean, rpos As Vector2i, objPos As Vector3i, tpos As Vector3i, canFire As Boolean

    actionUpdateFunc = UPDATE_FUNC_NONE

    checkDroid psDroid, "src/action.cpp:690", "actionUpdateDroid", max_check_object_recursion

    psPropStats = DROID_getPropulsionStats(psDroid)
    If Not (psPropStats <> -1) Then ASSERT_FAILURE "psPropStats != nullptr", "src/action.cpp:693", "actionUpdateDroid", "Invalid propulsion stats pointer": Exit Sub

    secHoldActive = secondaryGetState(psDroid, DSO_HALTTYPE) = DSS_HALT_HOLD

    actionSanity psDroid

    'if the droid has been attacked by an EMP weapon, it is temporarily disabled
    If aDroid(psDroid).lastHitWeapon = WSC_EMP Then
        If ULT(EMP_DISABLE_TIME, FromU32(U32(gameTime) - U32(aDroid(psDroid).timeLastHit))) Then
            'the actionStarted time needs to be adjusted
            aDroid(psDroid).actionStarted = FromU32(U32(aDroid(psDroid).actionStarted) + U32(FromU32(U32(gameTime) - U32(aDroid(psDroid).timeLastHit))))
            'reset the lastHit parameters
            aDroid(psDroid).timeLastHit = 0
            aDroid(psDroid).lastHitWeapon = WSC_NUM_WEAPON_SUBCLASSES
        Else
            'get out without updating
            Exit Sub
        End If
    End If

    For i = 0 To aDroid(psDroid).numWeaps - 1
        If aDroid(psDroid).asWeaps(i).nStat > 0 Then
            nonNullWeapon(i) = True
        End If
    Next

    ' HACK: Apparently we can't deal with a droid that only has NULL weapons ?
    ' FIXME: Find out whether this is really necessary
    If aDroid(psDroid).numWeaps <= 1 Then
        nonNullWeapon(0) = True
    End If

    Select Case aDroid(psDroid).action
    Case DACTION_NONE, DACTION_WAITFORREPAIR
        ' doing nothing
        ' see if there's anything to shoot.
        If aDroid(psDroid).numWeaps > 0 And Not DROID_isVtol(psDroid) _
           And (aDroid(psDroid).order.typ = DORDER_NONE Or aDroid(psDroid).order.typ = DORDER_HOLD Or aDroid(psDroid).order.typ = DORDER_RTR Or aDroid(psDroid).order.typ = DORDER_GUARD) Then
            For i = 0 To aDroid(psDroid).numWeaps - 1
                If nonNullWeapon(i) Then
                    psTemp = 0

                    psWeapStats = DROID_getWeaponStats(psDroid, i)
                    If aDroid(psDroid).asWeaps(i).nStat > 0 Then
                        If asWeaponStats(psWeapStats).rotate <> 0 Then
                            If IS_TIME_TO_CHECK_FOR_NEW_TARGET(psDroid) Then
                                If secondaryGetState(psDroid, DSO_ATTACK_LEVEL) = DSS_ALEV_ALWAYS Then
                                    ' don't bother doing this costly calculation again if aiUpdateDroid already checked this tick and failed
                                    If aDroid(psDroid).lastCheckNearestTargetFailed(i) <> gameTime Then
                                        If aiBestNearestTarget(psDroid, psTemp, i) >= 0 Then
                                            aDroid(psDroid).action = DACTION_ATTACK
                                            setDroidActionTarget psDroid, psTemp, i
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Case DACTION_WAITDURINGREPAIR
        ' don't want to be in a formation for this move
        If aDroid(psDroid).sMove.psFormation <> 0 Then
            formationLeave aDroid(psDroid).sMove.psFormation, psDroid
            aDroid(psDroid).sMove.psFormation = 0
        End If
        ' Check that repair facility still exists
        If aDroid(psDroid).order.psObj = 0 Then
            aDroid(psDroid).action = DACTION_NONE
            GoTo switchEnd
        End If
        If aDroid(psDroid).order.typ = DORDER_RTR And aDroid(psDroid).order.rtrType = RTR_TYPE_REPAIR_FACILITY Then
            ' move back to the repair facility if necessary
            If DROID_STOPPED(psDroid) Then
                objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj)
                If Not actionReachedBuildPos(psDroid, _
                                             objPos.x, objPos.y, aStruct(aDroid(psDroid).order.psObj).rot.direction, _
                                             structureStatsBase(aStruct(aDroid(psDroid).order.psObj).pStructureType)) Then
                    moveDroidToNoFormation psDroid, objPos.x, objPos.y
                End If
            End If
        ElseIf aDroid(psDroid).order.typ = DORDER_RTR And aDroid(psDroid).order.rtrType = RTR_TYPE_DROID And DROID_STOPPED(psDroid) Then
            If Not actionReachedDroid(psDroid, aDroid(psDroid).order.psObj) Then
                objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj)
                moveDroidToNoFormation psDroid, objPos.x, objPos.y
            Else
                moveStopDroid psDroid
            End If
        End If
    Case DACTION_TRANSPORTWAITTOFLYIN
        'if we're moving droids to safety and currently waiting to fly back in, see if time is up
        If aDroid(psDroid).player = selectedPlayer Then
            If getDroidsToSafetyFlag() Then
                enoughTimeRemaining = (CDbl(mission_time) - (U32(gameTime) - U32(mission_startTime))) >= (60 * GAME_TICKS_PER_SEC)
                If FromU32(U32(mission_ETA) - (U32(gameTime) - U32(missionGetReinforcementTime()))) <= 0 And enoughTimeRemaining Then
                    If Not droidRemove(psDroid, mission_apsDroidLists) Then
                        ASSERT_FAILURE "false", "src/action.cpp:809", "actionUpdateDroid", "Unable to remove transporter from mission list": Exit Sub
                    End If
                    addDroid psDroid, apsDroidLists
                    'set the x/y up since they were set to INVALID_XY when moved offWorld
                    missionGetTransporterExit selectedPlayer, droidX, droidY
                    aDroid(psDroid).pos.x = droidX
                    aDroid(psDroid).pos.y = droidY
                    'fly Transporter back to get some more droids
                    orderDroidLoc psDroid, DORDER_TRANSPORTIN, _
                                  getLandingX(selectedPlayer), getLandingY(selectedPlayer), ModeImmediate
                End If
            End If
        End If

    Case DACTION_MOVE, DACTION_RETURNTOPOS, DACTION_FIRESUPPORT_RETREAT
        ' moving to a location
        If DROID_STOPPED(psDroid) Then
            notify = aDroid(psDroid).action = DACTION_MOVE
            ' Got to destination
            aDroid(psDroid).action = DACTION_NONE

            If notify Then
                ' notify scripts we have reached the destination
                ' also triggers when patrolling and reached a waypoint

                triggerEventDroidIdle psDroid
            End If
        'added multiple weapon check
        ElseIf aDroid(psDroid).numWeaps > 0 Then
            For i = 0 To aDroid(psDroid).numWeaps - 1
                If nonNullWeapon(i) Then
                    psTemp = 0

                    'I moved psWeapStats flag update there
                    psWeapStats = DROID_getWeaponStats(psDroid, i)
                    If Not DROID_isVtol(psDroid) Then
                        If aDroid(psDroid).asWeaps(i).nStat > 0 Then
                            If asWeaponStats(psWeapStats).rotate <> 0 Then
                                If asWeaponStats(psWeapStats).fireOnMove Then
                                    If IS_TIME_TO_CHECK_FOR_NEW_TARGET(psDroid) Then
                                        If secondaryGetState(psDroid, DSO_ATTACK_LEVEL) = DSS_ALEV_ALWAYS Then
                                            If aiBestNearestTarget(psDroid, psTemp, i) >= 0 Then
                                                aDroid(psDroid).action = DACTION_MOVEFIRE
                                                setDroidActionTarget psDroid, psTemp, i
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Case DACTION_TRANSPORTIN, DACTION_TRANSPORTOUT
        actionUpdateTransporter psDroid
    Case DACTION_MOVEFIRE
        ' check if vtol is armed
        If vtolEmpty(psDroid) Then
            moveToRearm psDroid
        End If
        ' If droid stopped, it can no longer be in DACTION_MOVEFIRE
        If DROID_STOPPED(psDroid) Then
            aDroid(psDroid).action = DACTION_NONE
            GoTo switchEnd
        End If
        ' loop through weapons and look for target for each weapon
        bHasTarget = False
        For i = 0 To aDroid(psDroid).numWeaps - 1
            bDirect = proj_Direct(DROID_getWeaponStats(psDroid, i))
            blockingWall = 0
            ' Does this weapon have a target?
            If aDroid(psDroid).psActionTarget(i) <> 0 Then
                ' Is target worth shooting yet?
                If aiObjectIsProbablyDoomed(aDroid(psDroid).psActionTarget(i), bDirect) Then
                    setDroidActionTarget psDroid, 0, i
                ' Is target from our team now? (Electronic Warfare)
                ElseIf electronicDroidSameOwner(psDroid, i) Then
                    setDroidActionTarget psDroid, 0, i
                ' Is target blocked by a wall?
                ElseIf visGetBlockingWall_(bDirect, psDroid, aDroid(psDroid).psActionTarget(i)) Then
                    setDroidActionTarget psDroid, 0, i
                ' I have a target!
                Else
                    bHasTarget = True
                End If
            ' This weapon doesn't have a target
            Else
                ' Can we find a good target for the weapon?
                If aiBestNearestTarget(psDroid, psTemp, i) >= 0 Then ' assuming aiBestNearestTarget checks for electronic warfare
                    bHasTarget = True
                    setDroidActionTarget psDroid, psTemp, i ' this updates psDroid->psActionTarget[i] to != NULL
                End If
            End If
            ' If we have a target for the weapon: is it visible?
            If aDroid(psDroid).psActionTarget(i) <> 0 Then
                If visibleObject(psDroid, aDroid(psDroid).psActionTarget(i), False) > UBYTE_MAX \ 2 Then
                    hasVisibleTarget = True ' droid have a visible target to shoot
                    targetVisibile(i) = True ' it is at least visible for this weapon
                End If
            End If
        Next
        ' if there is at least one target
        If bHasTarget Then
            ' loop through weapons
            For i = 0 To aDroid(psDroid).numWeaps - 1
                compIndex = aDroid(psDroid).asWeaps(i).nStat
                wallBlocked = False

                ' has weapon a target? is target valid?
                If aDroid(psDroid).psActionTarget(i) <> 0 Then
                    If validTarget(psDroid, aDroid(psDroid).psActionTarget(i), i) Then
                        ' is target visible and weapon is not a Nullweapon?
                        If targetVisibile(i) And nonNullWeapon(i) Then 'to fix a AA-weapon attack ground unit exploit
                            psActionTarget = 0
                            blockingWall = visGetBlockingWall(psDroid, aDroid(psDroid).psActionTarget(i))

                            If proj_Direct(compIndex) And blockingWall <> 0 Then
                                weapEffect = asWeaponStats(compIndex).weaponEffect

                                If wallIsWorthShooting(psDroid, blockingWall, weapEffect) Then
                                    psActionTarget = blockingWall
                                    setDroidActionTarget psDroid, psActionTarget, i ' attack enemy wall
                                Else
                                    wallBlocked = True
                                End If
                            Else
                                psActionTarget = aDroid(psDroid).psActionTarget(i)
                            End If

                            ' is the turret aligned with the target?
                            If Not wallBlocked Then
                                If actionTargetTurret(psDroid, psActionTarget, i) Then
                                    ' In range - fire !!!
                                    combFire aDroid(psDroid).asWeaps(i), psDroid, psActionTarget, i
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            ' Droid don't have a visible target and it is not in pursue mode
            If Not hasVisibleTarget And secondaryGetState(psDroid, DSO_ATTACK_LEVEL) <> DSS_ALEV_ALWAYS Then
                ' Target lost
                aDroid(psDroid).action = DACTION_MOVE
            End If
        ' it don't have a target, change to DACTION_MOVE
        Else
            aDroid(psDroid).action = DACTION_MOVE
        End If
        'check its a VTOL unit since adding Transporter's into multiPlayer
        ' check vtol attack runs
        If DROID_isVtol(psDroid) Then
            actionUpdateVtolAttack psDroid
        End If
    Case DACTION_ATTACK, DACTION_ROTATETOATTACK
        If aDroid(psDroid).psActionTarget(0) = 0 And aDroid(psDroid).psActionTarget(1) <> 0 Then
            GoTo switchEnd
        End If

        If aDroid(psDroid).psActionTarget(0) = 0 Then
            If aDroid(psDroid).numWeaps > 1 Then
                ' FUTURE TODO: Figure out what (if any) changes to this switch case might make sense for multi-weapon bodies
                ' To avoid log spam, make this a silent return for now
                Exit Sub
            Else
                ' For single-weapon bodies, ASSERT
                If Not (aDroid(psDroid).psActionTarget(0) <> 0) Then ASSERT_FAILURE "psDroid->psActionTarget[0] != nullptr", "src/action.cpp:1017", "actionUpdateDroid", "target is NULL while attacking": Exit Sub
            End If
        End If

        If aDroid(psDroid).action = DACTION_ROTATETOATTACK Then
            If aDroid(psDroid).sMove.Status = MOVETURNTOTARGET Then
                objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
                moveTurnDroid psDroid, objPos.x, objPos.y
                GoTo switchEnd  ' Still turning.
            End If
            aDroid(psDroid).action = DACTION_ATTACK
        End If

        If aDroid(psDroid).action = DACTION_ATTACK Then
            ' don't wan't formations for this one
            If aDroid(psDroid).sMove.psFormation <> 0 Then
                formationLeave aDroid(psDroid).sMove.psFormation, psDroid
                aDroid(psDroid).sMove.psFormation = 0
            End If
        End If

        'check the target hasn't become one the same player ID - Electronic Warfare
        If electronicDroid(psDroid) Then
            If aDroid(psDroid).player = SIMPLE_OBJECT_player(aDroid(psDroid).psActionTarget(0)) Then
                For i = 0 To aDroid(psDroid).numWeaps - 1
                    setDroidActionTarget psDroid, 0, i
                Next
                aDroid(psDroid).action = DACTION_NONE
                GoTo switchEnd
            End If
        End If

        bHasTarget = False
        wallBlocked = False
        For i = 0 To aDroid(psDroid).numWeaps - 1
            If i > 0 Then
                ' If we're ordered to shoot something, and we can, shoot it
                If attackOrderCanShoot(psDroid, i) Then
                    setDroidActionTarget psDroid, aDroid(psDroid).psActionTarget(0), i
                ' If we still don't have a target, try to find one
                Else
                    If aDroid(psDroid).psActionTarget(i) = 0 Then
                        If IS_TIME_TO_CHECK_FOR_NEW_TARGET(psDroid) Then
                            If aiChooseTarget(psDroid, psTargets(i), i, False) Then ' Can probably just use psTarget instead of psTargets[i], and delete the psTargets variable.
                                setDroidActionTarget psDroid, psTargets(i), i
                            End If
                        End If
                    End If
                End If
            End If

            If aDroid(psDroid).psActionTarget(i) <> 0 Then
                psActionTarget = aDroid(psDroid).psActionTarget(i)
            Else
                psActionTarget = aDroid(psDroid).psActionTarget(0)
            End If

            canFire = False
            If nonNullWeapon(i) Then canFire = actionVisibleTarget_(psDroid, psActionTarget, i)
            If canFire Then
                psWeapStats = DROID_getWeaponStats(psDroid, i)
                weapEffect = asWeaponStats(psWeapStats).weaponEffect
                blockingWall = visGetBlockingWall(psDroid, psActionTarget)

                ' if a wall is inbetween us and the target, try firing at the wall if our
                ' weapon is good enough
                If proj_Direct(psWeapStats) And blockingWall <> 0 Then
                    If wallIsWorthShooting(psDroid, blockingWall, weapEffect) Then
                        psActionTarget = blockingWall
                        setDroidActionTarget psDroid, psActionTarget, i
                    Else
                        wallBlocked = True
                    End If
                End If

                If Not bHasTarget Then
                    bHasTarget = actionInRange(psDroid, psActionTarget, i, False)
                End If

                If validTarget(psDroid, psActionTarget, i) And Not wallBlocked Then
                    dirDiff = 0

                    If asWeaponStats(psWeapStats).rotate = 0 Then
                        ' no rotating turret - need to check aligned with target
                        objPos = SIMPLE_OBJECT_pos(psActionTarget)
                        targetDir = calcDirection(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, objPos.x, objPos.y)
                        dirDiff = Abs(angleDelta(targetDir - aDroid(psDroid).rot.direction))
                    End If

                    If dirDiff > FIXED_TURRET_DIR Then
                        If i > 0 Then
                            If aDroid(psDroid).psActionTarget(i) <> aDroid(psDroid).psActionTarget(0) Then
                                ' Nope, can't shoot this, try something else next time
                                setDroidActionTarget psDroid, 0, i
                            End If
                        ElseIf aDroid(psDroid).sMove.Status <> MOVESHUFFLE Then
                            aDroid(psDroid).action = DACTION_ROTATETOATTACK
                            objPos = SIMPLE_OBJECT_pos(psActionTarget)
                            moveTurnDroid psDroid, objPos.x, objPos.y
                        End If
                    ElseIf turretAlignedOrFixed(psDroid, psActionTarget, psWeapStats, i) Then
                        ' Prevents artillery units attached to a leader sometimes not stopping.
                        If Not DROID_isVtol(psDroid) And _
                           Not proj_Direct(psWeapStats) And _
                           aDroid(psDroid).order.typ = DORDER_FIRESUPPORT And _
                           Not DROID_STOPPED(psDroid) Then
                            moveStopDroid psDroid
                        End If
                        ' In range - fire !!!
                        combFire aDroid(psDroid).asWeaps(i), psDroid, psActionTarget, i
                    End If
                ElseIf i > 0 Then
                    ' Nope, can't shoot this, try something else next time
                    setDroidActionTarget psDroid, 0, i
                End If
            ElseIf i > 0 Then
                ' Nope, can't shoot this, try something else next time
                setDroidActionTarget psDroid, 0, i
            End If
        Next

        If Not bHasTarget Or wallBlocked Then
            supportsSensorTower = False
            If Not DROID_isVtol(psDroid) Then
                psTarget = orderStateObj(psDroid, DORDER_FIRESUPPORT)
                If psTarget <> 0 Then
                    supportsSensorTower = SIMPLE_OBJECT_type(psTarget) = OBJ_STRUCTURE
                End If
            End If

            If secHoldActive And (aDroid(psDroid).order.typ = DORDER_ATTACKTARGET Or aDroid(psDroid).order.typ = DORDER_FIRESUPPORT) Then
                aDroid(psDroid).action = DACTION_NONE ' secondary holding, cancel the order.
            ElseIf secondaryGetState(psDroid, DSO_HALTTYPE) = DSS_HALT_PURSUE And _
                   Not supportsSensorTower And _
                   Not (aDroid(psDroid).order.typ = DORDER_HOLD Or _
                        aDroid(psDroid).order.typ = DORDER_RTR) Then
                'We need this so pursuing doesn't stop if a unit is ordered to move somewhere while
                'it is still in weapon range of the target when reaching the end destination.
                'Weird case, I know, but keeps the previous pursue order intact.
                aDroid(psDroid).action = DACTION_MOVETOATTACK ' out of range - chase it
            ElseIf supportsSensorTower Or _
                   aDroid(psDroid).order.typ = DORDER_NONE Or _
                   aDroid(psDroid).order.typ = DORDER_HOLD Or _
                   aDroid(psDroid).order.typ = DORDER_RTR Then
                ' don't move if on hold or firesupport for a sensor tower
                ' also don't move if we're holding position or waiting for repair
                aDroid(psDroid).action = DACTION_NONE ' holding, cancel the order.
            'Units attached to commanders are always guarding the commander
            ElseIf secHoldActive And aDroid(psDroid).order.typ = DORDER_GUARD And hasCommander(psDroid) Then
                commander = aGroup(aDroid(psDroid).psGroup).psCommander

                If aDroid(commander).order.typ = DORDER_ATTACKTARGET Or _
                   aDroid(commander).order.typ = DORDER_FIRESUPPORT Or _
                   aDroid(commander).order.typ = DORDER_ATTACK Then
                    aDroid(psDroid).action = DACTION_MOVETOATTACK
                Else
                    aDroid(psDroid).action = DACTION_NONE
                End If
            ElseIf secondaryGetState(psDroid, DSO_HALTTYPE) <> DSS_HALT_HOLD Then
                aDroid(psDroid).action = DACTION_MOVETOATTACK ' out of range - chase it
            Else
                aDroid(psDroid).order.psObj = 0
                aDroid(psDroid).action = DACTION_NONE
            End If
        End If

    Case DACTION_VTOLATTACK
        psWeapStats = -1
        targetIsValid = validTarget(psDroid, aDroid(psDroid).psActionTarget(0), 0)
        'uses vtResult
        If aDroid(psDroid).psActionTarget(0) <> 0 And _
           targetIsValid Then
            'check if vtol that its armed
            If vtolEmptyOrLost(psDroid, targetIsValid) Then
                moveToRearm psDroid
                GoTo switchEnd
            End If

            For i = 0 To aDroid(psDroid).numWeaps - 1
                If nonNullWeapon(i) Then
                    If validTarget(psDroid, aDroid(psDroid).psActionTarget(0), i) Then
                        'I moved psWeapStats flag update there
                        psWeapStats = DROID_getWeaponStats(psDroid, i)
                        If actionVisibleTarget(psDroid, aDroid(psDroid).psActionTarget(0), i) Then
                            If actionInRange(psDroid, aDroid(psDroid).psActionTarget(0), i) Then
                                If aDroid(psDroid).player = selectedPlayer Then
                                    audio_QueueTrackMinDelay ID_SOUND_COMMENCING_ATTACK_RUN2, _
                                                             VTOL_ATTACK_AUDIO_DELAY
                                End If

                                If actionTargetTurret(psDroid, aDroid(psDroid).psActionTarget(0), i) Then
                                    ' In range - fire !!!
                                    combFire aDroid(psDroid).asWeaps(i), psDroid, _
                                             aDroid(psDroid).psActionTarget(0), i
                                End If
                            Else
                                actionTargetTurret psDroid, aDroid(psDroid).psActionTarget(0), i
                            End If
                        End If
                    End If
                End If
            Next
        End If

        ' circle around target if hovering and not cyborg
        attackRunDelta.x = aDroid(psDroid).pos.x - aDroid(psDroid).sMove.destination.x
        attackRunDelta.y = aDroid(psDroid).pos.y - aDroid(psDroid).sMove.destination.y
        If DROID_STOPPED(psDroid) Then
            actionAddVtolAttackRun psDroid
        ElseIf dot2i(attackRunDelta, attackRunDelta) < TILE_UNITS * TILE_UNITS Then
            actionAddVtolAttackRun psDroid
        ElseIf aDroid(psDroid).psActionTarget(0) <> 0 And _
               targetIsValid Then
            ' if the vtol is close to the target, go around again
            tpos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
            diff.x = aDroid(psDroid).pos.x - tpos.x
            diff.y = aDroid(psDroid).pos.y - tpos.y
            rangeSq = dot2i(diff, diff)
            If ULT(rangeSq, VTOL_ATTACK_TARDIST * VTOL_ATTACK_TARDIST) Then
                ' don't do another attack run if already moving away from the target
                diff.x = aDroid(psDroid).sMove.destination.x - tpos.x
                diff.y = aDroid(psDroid).sMove.destination.y - tpos.y
                If dot2i(diff, diff) < VTOL_ATTACK_TARDIST * VTOL_ATTACK_TARDIST Then
                    actionAddVtolAttackRun psDroid
                End If
            ' in case psWeapStats is still NULL
            ElseIf psWeapStats <> -1 Then
                ' if the vtol is far enough away head for the target again
                maxRange = proj_GetLongRange(asWeaponStats(psWeapStats), aDroid(psDroid).player)
                If ULT(MUL32(maxRange, maxRange), rangeSq) Then
                    ' don't do another attack run if already heading for the target
                    diff.x = aDroid(psDroid).sMove.destination.x - tpos.x
                    diff.y = aDroid(psDroid).sMove.destination.y - tpos.y
                    If dot2i(diff, diff) > VTOL_ATTACK_TARDIST * VTOL_ATTACK_TARDIST Then
                        moveDroidToDirect psDroid, tpos.x, tpos.y
                    End If
                End If
            End If
        End If
    Case DACTION_MOVETOATTACK
        ' don't wan't formations for this one
        If aDroid(psDroid).sMove.psFormation <> 0 Then
            formationLeave aDroid(psDroid).sMove.psFormation, psDroid
            aDroid(psDroid).sMove.psFormation = 0
        End If

        ' send vtols back to rearm
        If DROID_isVtol(psDroid) Then
            If vtolEmpty(psDroid) Then
                moveToRearm psDroid
                GoTo switchEnd
            End If
        End If

        If Not (aDroid(psDroid).psActionTarget(0) <> 0) Then ASSERT_FAILURE "psDroid->psActionTarget[0] != nullptr", "src/action.cpp:1337", "actionUpdateDroid", "action update move to attack target is NULL": Exit Sub
        For i = 0 To aDroid(psDroid).numWeaps - 1
            hasValidWeapon = hasValidWeapon Or validTarget(psDroid, aDroid(psDroid).psActionTarget(0), i)
        Next
        'check the target hasn't become one the same player ID - Electronic Warfare, and that the target is still valid.
        If electronicDroidSameOwner(psDroid) Or Not hasValidWeapon Then
            For i = 0 To aDroid(psDroid).numWeaps - 1
                setDroidActionTarget psDroid, 0, i
            Next
            aDroid(psDroid).action = DACTION_NONE
        Else
            If actionVisibleTarget(psDroid, aDroid(psDroid).psActionTarget(0), 0) Then
                For i = 0 To aDroid(psDroid).numWeaps - 1
                    If nonNullWeapon(i) Then
                        If validTarget(psDroid, aDroid(psDroid).psActionTarget(0), i) Then
                            If actionVisibleTarget(psDroid, aDroid(psDroid).psActionTarget(0), i) Then
                                chaseBloke = False
                                psWeapStats = DROID_getWeaponStats(psDroid, i)
                                actionIsInRange = actionInRange(psDroid, aDroid(psDroid).psActionTarget(0), i)
                                If asWeaponStats(psWeapStats).rotate <> 0 Then
                                    actionTargetTurret psDroid, aDroid(psDroid).psActionTarget(0), i
                                End If

                                If Not DROID_isVtol(psDroid) Then
                                    If SIMPLE_OBJECT_type(aDroid(psDroid).psActionTarget(0)) = OBJ_DROID Then
                                        If aDroid(aDroid(psDroid).psActionTarget(0)).droidType = DROID_PERSON And _
                                           asWeaponStats(psWeapStats).fireOnMove Then
                                            chaseBloke = True
                                        End If
                                    End If
                                End If

                                If actionIsInRange Then
                                    ' init vtol attack runs count if necessary
                                    If asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_LIFT Then
                                        aDroid(psDroid).action = DACTION_VTOLATTACK
                                    ElseIf chaseBloke Or Not actionInRange(psDroid, aDroid(psDroid).psActionTarget(0), i, False) Then
                                        ' fire while closing range
                                        ' do this either when trying to run over a bloke or close the distance into optimum range
                                        blockingWall = visGetBlockingWall(psDroid, aDroid(psDroid).psActionTarget(0))
                                        If blockingWall <> 0 And proj_Direct(psWeapStats) Then
                                            weapEffect = asWeaponStats(psWeapStats).weaponEffect

                                            If wallIsWorthShooting(psDroid, blockingWall, weapEffect) Then
                                                'Shoot at wall if the weapon is good enough against them
                                                combFire aDroid(psDroid).asWeaps(i), psDroid, blockingWall, i
                                            End If
                                        Else
                                            combFire aDroid(psDroid).asWeaps(i), psDroid, aDroid(psDroid).psActionTarget(0), i
                                        End If
                                    Else
                                        ' stop and shoot
                                        moveStopDroid psDroid

                                        If asWeaponStats(psWeapStats).rotate <> 0 Then
                                            aDroid(psDroid).action = DACTION_ATTACK
                                        Else
                                            aDroid(psDroid).action = DACTION_ROTATETOATTACK
                                            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
                                            moveTurnDroid psDroid, objPos.x, objPos.y
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            Else
                For i = 0 To aDroid(psDroid).numWeaps - 1
                    If (aDroid(psDroid).asWeaps(i).rot.direction <> 0) Or _
                       (aDroid(psDroid).asWeaps(i).rot.pitch <> 0) Then
                        actionAlignTurret psDroid, i
                    End If
                Next
            End If

            If DROID_STOPPED(psDroid) And aDroid(psDroid).action <> DACTION_ATTACK Then
                ' Stopped moving but haven't reached the target - possibly move again

                ''hack' to make the droid to check the primary turrent instead of all
                psWeapStats = DROID_getWeaponStats(psDroid, 0)

                If aDroid(psDroid).order.typ = DORDER_ATTACKTARGET And secHoldActive Then
                    aDroid(psDroid).action = DACTION_NONE ' on hold, give up.
                ElseIf actionInsideMinRange(psDroid, aDroid(psDroid).psActionTarget(0), psWeapStats) Then
                    If proj_Direct(psWeapStats) And aDroid(psDroid).order.typ <> DORDER_HOLD Then
                        ' try and extend the range
                        actionCalcPullBackPoint psDroid, aDroid(psDroid).psActionTarget(0), pbx, pby
                        moveDroidTo psDroid, pbx, pby
                    Else
                        If asWeaponStats(psWeapStats).rotate <> 0 Then
                            aDroid(psDroid).action = DACTION_ATTACK
                        Else
                            aDroid(psDroid).action = DACTION_ROTATETOATTACK
                            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
                            moveTurnDroid psDroid, objPos.x, _
                                          objPos.y
                        End If
                    End If
                ElseIf aDroid(psDroid).order.typ <> DORDER_HOLD Then ' approach closer?
                    ' try to close the range
                    objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
                    moveDroidTo psDroid, objPos.x, objPos.y
                End If
            End If
        End If

    Case DACTION_SULK
        ' unable to route to target ... don't do anything aggressive until time is up
        ' we need to do something defensive at this point ???

        'hmmm, hope this doesn't cause any problems!
        If ULT(aDroid(psDroid).actionStarted, gameTime) Then
            aDroid(psDroid).action = DACTION_NONE ' Sulking is over lets get back to the action ... is this all I need to do to get it back into the action?
        End If

    Case DACTION_MOVETOBUILD
        If aDroid(psDroid).order.psStats = -1 Then
            aDroid(psDroid).action = DACTION_NONE
            GoTo switchEnd
        Else
            ' Determine if the droid can still build or help to build the ordered structure at the specified location
            desiredStructure = aDroid(psDroid).order.psStats
            structureAtBuildPosition = getTileStructure(map_coord(aDroid(psDroid).actionPos.x), map_coord(aDroid(psDroid).actionPos.y))

            If 0 <> structureAtBuildPosition Then
                droidCannotBuild = False

                If Not aiCheckAlliances(aStruct(structureAtBuildPosition).player, aDroid(psDroid).player) Then
                    ' Not our structure
                    droidCannotBuild = True
                ' There's an allied structure already there.  Is it a wall, and can the droid upgrade it to a defence or gate?
                ElseIf isWall(asStructureStats(aStruct(structureAtBuildPosition).pStructureType).typ) And _
                       (asStructureStats(desiredStructure).typ = REF_DEFENSE Or asStructureStats(desiredStructure).typ = REF_GATE) Then
                    ' It's always valid to upgrade a wall to a defence or gate
                    droidCannotBuild = False ' Just to avoid an empty branch
                ElseIf (aStruct(structureAtBuildPosition).pStructureType <> desiredStructure) And _
                       (asStructureStats(aStruct(structureAtBuildPosition).pStructureType).typ = REF_WALLCORNER And asStructureStats(desiredStructure).typ <> REF_WALL) Then ' and not a wall corner when the droid wants to build a wall
                    ' And so the droid can't build or help with building this structure
                    droidCannotBuild = True
                ' So it's a structure that the droid could help to build, but is it already complete?
                ElseIf aStruct(structureAtBuildPosition).status = SS_BUILT Then
                    If Not IsStatExpansionModule(desiredStructure) Then
                        ' The building is complete and the droid hasn't been told to add a module, or can't add one, so can't help with that.
                        droidCannotBuild = True
                    ElseIf Not canStructureHaveAModuleAdded(structureAtBuildPosition) Then
                        droidCannotBuild = True
                    End If
                End If

                If droidCannotBuild Then
                    If aDroid(psDroid).order.typ = DORDER_LINEBUILD And _
                       (map_coord(aDroid(psDroid).order.pos.x) <> map_coord(aDroid(psDroid).order.pos2.x) Or map_coord(aDroid(psDroid).order.pos.y) <> map_coord(aDroid(psDroid).order.pos2.y)) Then
                        ' The droid is doing a line build, and there's more to build. This will force the droid to move to the next structure in the line build
                        objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_MOVETOBUILD: line target is already built, or can't be built - moving to next structure in line"
                        aDroid(psDroid).action = DACTION_NONE
                    Else
                        ' Cancel the current build order. This will move the truck onto the next order, if it has one, or halt in place.
                        objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_MOVETOBUILD: target is already built, or can't be built - executing next order or halting"
                        cancelBuild psDroid
                    End If

                    GoTo switchEnd
                End If
            End If
        End If ' End of check for whether the droid can still succesfully build the ordered structure

        ' The droid cannot be in a formation
        If aDroid(psDroid).sMove.psFormation <> 0 Then
            formationLeave aDroid(psDroid).sMove.psFormation, psDroid
            aDroid(psDroid).sMove.psFormation = 0
        End If

        ' The droid can still build or help with a build, and is moving to a location to do so - are we there yet, are we there yet, are we there yet?
        If actionReachedBuildPos(psDroid, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y, aDroid(psDroid).order.direction, structureStatsBase(aDroid(psDroid).order.psStats)) Then
            ' We're there, go ahead and build or help to build the structure
            buildPosEmpty = actionRemoveDroidsFromBuildPos(aDroid(psDroid).player, aDroid(psDroid).actionPos, aDroid(psDroid).order.direction, structureStatsBase(aDroid(psDroid).order.psStats))
            If Not buildPosEmpty Then
                GoTo switchEnd
            End If

            helpBuild = False
            ' Got to destination - start building
            psStructStats = aDroid(psDroid).order.psStats
            dir_ = aDroid(psDroid).order.direction
            moveStopDroid psDroid
            objTrace aDroid(psDroid).id, "actionUpdateDroid", "Halted in our tracks - at construction site"
            If aDroid(psDroid).order.typ = DORDER_BUILD And aDroid(psDroid).order.psObj = 0 Then
                ' Starting a new structure
                bpos = aDroid(psDroid).actionPos

                'need to check if something has already started building here?
                'unless its a module!
                If IsStatExpansionModule(psStructStats) Then
                    syncDebug "actionUpdateDroid", "Reached build target: module"
                    debug_ 1588, LOG_NEVER, "actionUpdateDroid", "DACTION_MOVETOBUILD: setUpBuildModule"
                    setUpBuildModule psDroid
                ElseIf TileHasStructure(worldTile(bpos.x, bpos.y)) Then
                    ' structure on the build location - see if it is the same type
                    psStruct = getTileStructure(map_coord(bpos.x), map_coord(bpos.y))
                    If aStruct(psStruct).pStructureType = aDroid(psDroid).order.psStats Or _
                       (asStructureStats(aDroid(psDroid).order.psStats).typ = REF_WALL And asStructureStats(aStruct(psStruct).pStructureType).typ = REF_WALLCORNER) Then
                        ' same type - do a help build
                        syncDebug "actionUpdateDroid", "Reached build target: do-help"
                        setDroidTarget psDroid, psStruct
                        helpBuild = True
                    ElseIf (asStructureStats(aStruct(psStruct).pStructureType).typ = REF_WALL Or _
                            asStructureStats(aStruct(psStruct).pStructureType).typ = REF_WALLCORNER) And _
                           (asStructureStats(aDroid(psDroid).order.psStats).typ = REF_DEFENSE Or _
                            asStructureStats(aDroid(psDroid).order.psStats).typ = REF_GATE) Then
                        ' building a gun tower or gate over a wall - OK
                        If droidStartBuild(psDroid) <> DroidStartBuildFailed Then
                            syncDebug "actionUpdateDroid", "Reached build target: tower"
                            aDroid(psDroid).action = DACTION_BUILD
                        End If
                    Else
                        syncDebug "actionUpdateDroid", "Reached build target: already-structure"
                        objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_MOVETOBUILD: tile has structure already"
                        cancelBuild psDroid
                    End If
                ElseIf Not validLocation(structureStatsBase(aDroid(psDroid).order.psStats), bpos, dir_, aDroid(psDroid).player, False) Then
                    syncDebug "actionUpdateDroid", "Reached build target: invalid"
                    objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_MOVETOBUILD: !validLocation"
                    cancelBuild psDroid
                ElseIf droidStartBuild(psDroid) = DroidStartBuildSuccess Then  ' If DroidStartBuildPending, then there's a burning oil well, and we don't want to change to DACTION_BUILD until it stops burning.
                    syncDebug "actionUpdateDroid", "Reached build target: build"
                    aDroid(psDroid).action = DACTION_BUILD
                    aDroid(psDroid).actionStarted = gameTime
                    aDroid(psDroid).actionPoints = 0
                End If
            ElseIf aDroid(psDroid).order.typ = DORDER_LINEBUILD Or aDroid(psDroid).order.typ = DORDER_BUILD Then
                ' building a wall.
                psTile = worldTile(aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y)
                syncDebug "actionUpdateDroid", "Reached build target: wall"
                If aDroid(psDroid).order.psObj = 0 _
                   And (TileHasStructure(psTile) _
                        Or TileHasFeature(psTile)) Then
                    If TileHasStructure(psTile) Then
                        ' structure on the build location - see if it is the same type
                        psStruct = getTileStructure(map_coord(aDroid(psDroid).actionPos.x), map_coord(aDroid(psDroid).actionPos.y))
                        If Not (psStruct <> 0) Then ASSERT_FAILURE "psStruct", "src/action.cpp:1649", "actionUpdateDroid", "TileHasStructure, but getTileStructure returned nullptr"
                        If aStruct(psStruct).pStructureType = aDroid(psDroid).order.psStats Then
                            ' same type - do a help build
                            setDroidTarget psDroid, psStruct
                            helpBuild = True
                        ElseIf (asStructureStats(aStruct(psStruct).pStructureType).typ = REF_WALL Or asStructureStats(aStruct(psStruct).pStructureType).typ = REF_WALLCORNER) And _
                               (asStructureStats(aDroid(psDroid).order.psStats).typ = REF_DEFENSE Or asStructureStats(aDroid(psDroid).order.psStats).typ = REF_GATE) Then
                            ' building a gun tower over a wall - OK
                            If droidStartBuild(psDroid) <> DroidStartBuildFailed Then
                                objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_MOVETOBUILD: start building defense"
                                aDroid(psDroid).action = DACTION_BUILD
                            End If
                        ElseIf IsStatExpansionModule(psStructStats) Then
                            setUpBuildModule psDroid
                        Else
                            objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_MOVETOBUILD: line build hit building"
                            cancelBuild psDroid
                        End If
                    ElseIf TileHasFeature(psTile) Then
                        feature_ = getTileFeature(map_coord(aDroid(psDroid).actionPos.x), map_coord(aDroid(psDroid).actionPos.y))
                        objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_MOVETOBUILD: tile has feature %d", asFeatureStats(aFeature(feature_).psStats).subType
                        If asFeatureStats(aFeature(feature_).psStats).subType = FEAT_OIL_RESOURCE And asStructureStats(aDroid(psDroid).order.psStats).typ = REF_RESOURCE_EXTRACTOR Then
                            If droidStartBuild(psDroid) <> DroidStartBuildFailed Then
                                objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_MOVETOBUILD: start building oil derrick"
                                aDroid(psDroid).action = DACTION_BUILD
                            End If
                        End If
                    Else
                        objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_MOVETOBUILD: blocked line build"
                        cancelBuild psDroid
                    End If
                ElseIf droidStartBuild(psDroid) <> DroidStartBuildFailed Then
                    aDroid(psDroid).action = DACTION_BUILD
                End If
            Else
                syncDebug "actionUpdateDroid", "Reached build target: planned-help"
                objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_MOVETOBUILD: planned-help"
                helpBuild = True
            End If

            If helpBuild Then
                ' continuing a partially built structure (order = helpBuild)
                If droidStartBuild(psDroid) <> DroidStartBuildFailed Then
                    objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_MOVETOBUILD: starting help build"
                    aDroid(psDroid).action = DACTION_BUILD
                End If
            End If
        ElseIf DROID_STOPPED(psDroid) Then
            objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_MOVETOBUILD: Starting to drive toward construction site - move status was %d", aDroid(psDroid).sMove.Status
            moveDroidToNoFormation psDroid, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y
        End If
    Case DACTION_BUILD
        If aDroid(psDroid).order.psStats = -1 Then
            objTrace aDroid(psDroid).id, "actionUpdateDroid", "No target stats for build order - resetting"
            aDroid(psDroid).action = DACTION_NONE
            GoTo switchEnd
        End If
        ' The droid cannot be in a formation
        If aDroid(psDroid).sMove.psFormation <> 0 Then
            formationLeave aDroid(psDroid).sMove.psFormation, psDroid
            aDroid(psDroid).sMove.psFormation = 0
        End If
        If stoppedShortOfBuildPos(psDroid, 0, aDroid(psDroid).order.direction, structureStatsBase(aDroid(psDroid).order.psStats)) Then
            objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_BUILD: Starting to drive toward construction site"
            moveDroidToNoFormation psDroid, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y
        ElseIf Not DROID_STOPPED(psDroid) And _
               aDroid(psDroid).sMove.Status <> MOVETURNTOTARGET And _
               aDroid(psDroid).sMove.Status <> MOVESHUFFLE Then
            If actionReachedBuildPos(psDroid, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y, aDroid(psDroid).order.direction, structureStatsBase(aDroid(psDroid).order.psStats)) Then
                objTrace aDroid(psDroid).id, "actionUpdateDroid", "DACTION_BUILD: Stopped - at construction site"
                moveStopDroid psDroid
            End If
        End If
        If aDroid(psDroid).action = DACTION_SULK Then
            objTrace aDroid(psDroid).id, "actionUpdateDroid", "Failed to go to objective, aborting build action"
            aDroid(psDroid).action = DACTION_NONE
            GoTo switchEnd
        End If
        If droidUpdateBuild(psDroid) Then
            actionTargetTurret psDroid, aDroid(psDroid).psActionTarget(0), 0
        End If
    Case DACTION_MOVETODEMOLISH, DACTION_MOVETOREPAIR, DACTION_MOVETORESTORE
        If aDroid(psDroid).order.psStats = -1 Then
            aDroid(psDroid).action = DACTION_NONE
            GoTo switchEnd
        Else
            structureAtPos = getTileStructure(map_coord(aDroid(psDroid).actionPos.x), map_coord(aDroid(psDroid).actionPos.y))

            If structureAtPos = 0 Or aDroid(psDroid).psActionTarget(0) = 0 Then
                'No structure located at desired position. Move on.
                aDroid(psDroid).action = DACTION_NONE
                GoTo switchEnd
            ElseIf aDroid(psDroid).order.typ <> DORDER_RESTORE Then
                cantDoRepairLikeAction = False

                If Not aiCheckAlliances(aStruct(structureAtPos).player, aDroid(psDroid).player) Then
                    cantDoRepairLikeAction = True
                ElseIf aDroid(psDroid).order.typ <> DORDER_DEMOLISH And aStruct(structureAtPos).body = STRUCTURE_structureBody(structureAtPos) Then
                    cantDoRepairLikeAction = True
                ElseIf aDroid(psDroid).order.typ = DORDER_DEMOLISH And aStruct(structureAtPos).player <> aDroid(psDroid).player Then
                    ObjAllied = aiCheckAlliances(aDroid(psDroid).player, aStruct(structureAtPos).player)
                    ObjDepartedAlly = False
                    If bMultiPlayer Then
                        If NetPlay_players_difficulty(aStruct(structureAtPos).player) = AIDifficulty_HUMAN Then
                            ObjDepartedAlly = Not NetPlay_players_allocated(aStruct(structureAtPos).player)
                        End If
                    End If
                    If Not ObjAllied Or Not ObjDepartedAlly Then
                        cantDoRepairLikeAction = True
                    End If
                End If

                If cantDoRepairLikeAction Then
                    aDroid(psDroid).action = DACTION_NONE
                    moveStopDroid psDroid
                    GoTo switchEnd
                End If
            End If
        End If

        ' The droid cannot be in a formation
        If aDroid(psDroid).sMove.psFormation <> 0 Then
            formationLeave aDroid(psDroid).sMove.psFormation, psDroid
            aDroid(psDroid).sMove.psFormation = 0
        End If

        ' see if the droid is at the edge of what it is moving to
        If actionReachedBuildPos(psDroid, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y, aStruct(aDroid(psDroid).psActionTarget(0)).rot.direction, structureStatsBase(aDroid(psDroid).order.psStats)) Then
            moveStopDroid psDroid

            ' got to the edge - start doing whatever it was meant to do
            droidStartAction psDroid
            Select Case aDroid(psDroid).action
            Case DACTION_MOVETODEMOLISH
                aDroid(psDroid).action = DACTION_DEMOLISH
            Case DACTION_MOVETOREPAIR
                aDroid(psDroid).action = DACTION_REPAIR
            Case DACTION_MOVETORESTORE
                aDroid(psDroid).action = DACTION_RESTORE
            Case Else
            End Select
        ElseIf DROID_STOPPED(psDroid) Then
            moveDroidToNoFormation psDroid, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y
        End If

    Case DACTION_DEMOLISH, DACTION_REPAIR, DACTION_RESTORE
        If aDroid(psDroid).order.psStats = -1 Or aDroid(psDroid).psActionTarget(0) = 0 Then
            aDroid(psDroid).action = DACTION_NONE
            GoTo switchEnd
        End If
        ' set up for the specific action
        Select Case aDroid(psDroid).action
        Case DACTION_DEMOLISH
            ' DACTION_MOVETODEMOLISH;
            actionUpdateFunc = UPDATE_FUNC_DEMOLISHING
        Case DACTION_REPAIR
            ' DACTION_MOVETOREPAIR;
            actionUpdateFunc = UPDATE_FUNC_REPAIR
        Case DACTION_RESTORE
            ' DACTION_MOVETORESTORE;
            actionUpdateFunc = UPDATE_FUNC_RESTORE
        Case Else
        End Select

        ' now do the action update
        If stoppedShortOfBuildPos(psDroid, aDroid(psDroid).psActionTarget(0), 0, structureStatsBase(aDroid(psDroid).order.psStats)) Then
            If aDroid(psDroid).order.typ <> DORDER_HOLD And (Not secHoldActive Or (secHoldActive And aDroid(psDroid).order.typ <> DORDER_NONE)) Then
                objTrace aDroid(psDroid).id, "actionUpdateDroid", "Secondary order: Go to construction site"
                moveDroidToNoFormation psDroid, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y
            Else
                aDroid(psDroid).action = DACTION_NONE
            End If
        ElseIf movingAtBuildPos(psDroid, aDroid(psDroid).psActionTarget(0), structureStatsBase(aDroid(psDroid).order.psStats)) Then
            objTrace aDroid(psDroid).id, "actionUpdateDroid", "Stopped - reached build position"
            moveStopDroid psDroid
        ElseIf actionUpdateFunc_(actionUpdateFunc, psDroid) Then
            'use 0 for non-combat(only 1 'weapon')
            actionTargetTurret psDroid, aDroid(psDroid).psActionTarget(0), 0
        Else
            aDroid(psDroid).action = DACTION_NONE
        End If

    Case DACTION_MOVETOREARMPOINT
        If DROID_STOPPED(psDroid) Then
            objTrace aDroid(psDroid).id, "actionUpdateDroid", "Finished moving onto the rearm pad"
            aDroid(psDroid).action = DACTION_WAITDURINGREARM
        End If
    Case DACTION_MOVETOREPAIRPOINT
        ' don't want to be in a formation for this move
        If aDroid(psDroid).sMove.psFormation <> 0 Then
            formationLeave aDroid(psDroid).sMove.psFormation, psDroid
            aDroid(psDroid).sMove.psFormation = 0
        End If
        If aDroid(psDroid).order.rtrType = RTR_TYPE_REPAIR_FACILITY Then
            ' moving from front to rear of repair facility or rearm pad
            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
            If actionReachedBuildPos(psDroid, objPos.x, objPos.y, aStruct(aDroid(psDroid).psActionTarget(0)).rot.direction, structureStatsBase(aStruct(aDroid(psDroid).psActionTarget(0)).pStructureType)) Then
                objTrace aDroid(psDroid).id, "actionUpdateDroid", "Arrived at repair point - waiting for our turn"
                moveStopDroid psDroid
                aDroid(psDroid).action = DACTION_WAITDURINGREPAIR
            ElseIf DROID_STOPPED(psDroid) Then
                moveDroidToNoFormation psDroid, objPos.x, _
                                       objPos.y
            End If
        ElseIf aDroid(psDroid).order.rtrType = RTR_TYPE_DROID Then
            reached = actionReachedDroid(psDroid, aDroid(psDroid).order.psObj)
            If reached Then
                If aDroid(psDroid).body >= aDroid(psDroid).originalBody Then
                    objTrace aDroid(psDroid).id, "actionUpdateDroid", "Repair not needed of droid %d", aDroid(psDroid).id
                    ' set droid points to max
                    aDroid(psDroid).body = aDroid(psDroid).originalBody
                    ' if completely repaired then reset order
                    secondarySetState psDroid, DSO_RETURN_TO_LOC, DSS_NONE
                    orderDroidObj psDroid, DORDER_GUARD, aDroid(psDroid).order.psObj, ModeImmediate
                Else
                    objTrace aDroid(psDroid).id, "actionUpdateDroid", "Stopping and waiting for repairs %d", aDroid(psDroid).id
                    moveStopDroid psDroid
                    aDroid(psDroid).action = DACTION_WAITDURINGREPAIR
                End If

            ElseIf DROID_STOPPED(psDroid) Then
                'objTrace(psDroid->id, "Droid was stopped, but havent reach the target, moving now");
                'moveDroidToNoFormation(psDroid, psDroid->order.psObj->pos.x, psDroid->order.psObj->pos.y);
            End If
        End If
    Case DACTION_OBSERVE
        ' align the turret
        actionTargetTurret psDroid, aDroid(psDroid).psActionTarget(0), 0

        If Not cbSensorDroid(psDroid) And aDroid(psDroid).psActionTarget(0) <> 0 Then
            ' make sure the target is within sensor range
            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
            xdiff = aDroid(psDroid).pos.x - objPos.x
            ydiff = aDroid(psDroid).pos.y - objPos.y
            rangeSq = droidSensorRange(psDroid)
            rangeSq = rangeSq * rangeSq
            If visibleObject(psDroid, aDroid(psDroid).psActionTarget(0), False) = 0 _
               Or xdiff * xdiff + ydiff * ydiff >= rangeSq Then
                If secondaryGetState(psDroid, DSO_HALTTYPE) <> DSS_HALT_GUARD And (aDroid(psDroid).order.typ = DORDER_NONE Or aDroid(psDroid).order.typ = DORDER_HOLD) Then
                    aDroid(psDroid).action = DACTION_NONE
                ElseIf (Not secHoldActive And aDroid(psDroid).order.typ <> DORDER_HOLD) Or (secHoldActive And aDroid(psDroid).order.typ = DORDER_OBSERVE) Then
                    aDroid(psDroid).action = DACTION_MOVETOOBSERVE
                    moveDroidTo psDroid, objPos.x, objPos.y
                End If
            End If
        End If
    Case DACTION_MOVETOOBSERVE
        ' align the turret
        actionTargetTurret psDroid, aDroid(psDroid).psActionTarget(0), 0

        If visibleObject(psDroid, aDroid(psDroid).psActionTarget(0), False) <> 0 Then
            ' make sure the target is within sensor range
            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
            xdiff = aDroid(psDroid).pos.x - objPos.x
            ydiff = aDroid(psDroid).pos.y - objPos.y
            rangeSq = droidSensorRange(psDroid)
            rangeSq = rangeSq * rangeSq
            If (xdiff * xdiff + ydiff * ydiff < rangeSq) And _
               Not DROID_STOPPED(psDroid) Then
                aDroid(psDroid).action = DACTION_OBSERVE
                moveStopDroid psDroid
            End If
        End If
        If DROID_STOPPED(psDroid) And aDroid(psDroid).action = DACTION_MOVETOOBSERVE Then
            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
            moveDroidTo psDroid, objPos.x, objPos.y
        End If
    Case DACTION_FIRESUPPORT
        If aDroid(psDroid).order.psObj = 0 Then
            aDroid(psDroid).action = DACTION_NONE
            Exit Sub
        End If
        'can be either a droid or a structure now - AB 7/10/98
        If Not ((SIMPLE_OBJECT_type(aDroid(psDroid).order.psObj) = OBJ_DROID Or SIMPLE_OBJECT_type(aDroid(psDroid).order.psObj) = OBJ_STRUCTURE) _
                And aiCheckAlliances(SIMPLE_OBJECT_player(aDroid(psDroid).order.psObj), aDroid(psDroid).player)) Then ASSERT_FAILURE "(order->psObj->type == OBJ_DROID || order->psObj->type == OBJ_STRUCTURE) && aiCheckAlliances(order->psObj->player, psDroid->player)", "src/action.cpp:2024", "actionUpdateDroid", "DACTION_FIRESUPPORT: incorrect target type": Exit Sub

        'don't move VTOL's
        ' also don't move closer to sensor towers
        If Not DROID_isVtol(psDroid) And SIMPLE_OBJECT_type(aDroid(psDroid).order.psObj) <> OBJ_STRUCTURE Then
            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj)
            diff.x = aDroid(psDroid).pos.x - objPos.x
            diff.y = aDroid(psDroid).pos.y - objPos.y
            'Consider .shortRange here
            rangeSq = asWeaponStats(DROID_getWeaponStats(psDroid, 0)).upgrade(aDroid(psDroid).player).maxRange \ 2 ' move close to sensor
            rangeSq = rangeSq * rangeSq
            If dot2i(diff, diff) < rangeSq Then
                If Not DROID_STOPPED(psDroid) Then
                    moveStopDroid psDroid
                End If
            Else
                If Not DROID_STOPPED(psDroid) Then
                    diff.x = objPos.x - aDroid(psDroid).sMove.destination.x
                    diff.y = objPos.y - aDroid(psDroid).sMove.destination.y
                End If
                If DROID_STOPPED(psDroid) Or dot2i(diff, diff) > rangeSq Then
                    If secHoldActive Then
                        ' droid on hold, don't allow moves.
                        aDroid(psDroid).action = DACTION_NONE
                    Else
                        ' move in range
                        moveDroidTo psDroid, objPos.x, objPos.y
                    End If
                End If
            End If
        End If
    Case DACTION_MOVETODROIDREPAIR
        actionTargetObj = aDroid(psDroid).psActionTarget(0)
        If Not (actionTargetObj <> 0) Then ASSERT_FAILURE "actionTargetObj != nullptr && actionTargetObj->type == OBJ_DROID", "src/action.cpp:2067", "actionUpdateDroid", "unexpected repair target": Exit Sub
        If Not (SIMPLE_OBJECT_type(actionTargetObj) = OBJ_DROID) Then ASSERT_FAILURE "actionTargetObj != nullptr && actionTargetObj->type == OBJ_DROID", "src/action.cpp:2067", "actionUpdateDroid", "unexpected repair target": Exit Sub
        actionTarget = actionTargetObj
        If aDroid(actionTarget).body = aDroid(actionTarget).originalBody Then
            ' target is healthy: nothing to do
            aDroid(psDroid).action = DACTION_NONE
            moveStopDroid psDroid
            GoTo switchEnd
        End If
        objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
        diff.x = aDroid(psDroid).pos.x - objPos.x
        diff.y = aDroid(psDroid).pos.y - objPos.y
        ' moving to repair a droid
        If aDroid(psDroid).psActionTarget(0) = 0 Then  ' Target missing.
            aDroid(psDroid).action = DACTION_NONE
            Exit Sub
        ElseIf aDroid(psDroid).order.typ <> DORDER_DROIDREPAIR And dot2i(diff, diff) > 2 * REPAIR_MAXDIST * REPAIR_MAXDIST Then  ' Target farther then 1.4142 * REPAIR_MAXDIST and we aren't ordered to follow.
            aDroid(psDroid).action = DACTION_NONE
            Exit Sub
        End If
        If dot2i(diff, diff) < REPAIR_RANGE * REPAIR_RANGE Then
            ' Got to destination - start repair
            'rotate turret to point at droid being repaired
            'use 0 for repair droid
            actionTargetTurret psDroid, aDroid(psDroid).psActionTarget(0), 0
            droidStartAction psDroid
            aDroid(psDroid).action = DACTION_DROIDREPAIR
            droidRepairStarted aDroid(psDroid).psActionTarget(0), psDroid
        End If
        If DROID_STOPPED(psDroid) Then
            ' Couldn't reach destination - try and find a new one
            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
            aDroid(psDroid).actionPos.x = objPos.x
            aDroid(psDroid).actionPos.y = objPos.y
            moveDroidTo psDroid, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y
        End If
    Case DACTION_DROIDREPAIR
        If aDroid(psDroid).psActionTarget(0) = 0 Then
            aDroid(psDroid).action = DACTION_NONE
            GoTo switchEnd
        End If

        ' If not doing self-repair (psActionTarget[0] is repair target)
        If aDroid(psDroid).psActionTarget(0) <> psDroid Then
            actionTargetTurret psDroid, aDroid(psDroid).psActionTarget(0), 0
        ' Just self-repairing.
        ' See if there's anything to shoot.
        ElseIf aDroid(psDroid).numWeaps > 0 And Not DROID_isVtol(psDroid) _
               And (aDroid(psDroid).order.typ = DORDER_NONE Or aDroid(psDroid).order.typ = DORDER_HOLD Or aDroid(psDroid).order.typ = DORDER_RTR) Then
            For i = 0 To aDroid(psDroid).numWeaps - 1
                If nonNullWeapon(i) Then
                    psTemp = 0

                    psWeapStats = DROID_getWeaponStats(psDroid, i)
                    If aDroid(psDroid).asWeaps(i).nStat > 0 Then
                        If asWeaponStats(psWeapStats).rotate <> 0 Then
                            If secondaryGetState(psDroid, DSO_ATTACK_LEVEL) = DSS_ALEV_ALWAYS Then
                                If IS_TIME_TO_CHECK_FOR_NEW_TARGET(psDroid) Then
                                    If aiBestNearestTarget(psDroid, psTemp, i) >= 0 Then
                                        If psTemp <> 0 Then
                                            droidRepairStopped castDroid(aDroid(psDroid).psActionTarget(0)), psDroid
                                            aDroid(psDroid).action = DACTION_ATTACK
                                            setDroidActionTarget psDroid, psTemp, 0
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        If aDroid(psDroid).action <> DACTION_DROIDREPAIR Then
            GoTo switchEnd ' action has changed
        End If

        'check still next to the damaged droid
        objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
        xdiff = aDroid(psDroid).pos.x - objPos.x
        ydiff = aDroid(psDroid).pos.y - objPos.y
        If xdiff * xdiff + ydiff * ydiff > REPAIR_RANGE * REPAIR_RANGE Then
            droidRepairStopped castDroid(aDroid(psDroid).psActionTarget(0)), psDroid
            If aDroid(psDroid).order.typ = DORDER_DROIDREPAIR Then
                ' damaged droid has moved off - follow if we're not holding position!
                aDroid(psDroid).actionPos.x = objPos.x
                aDroid(psDroid).actionPos.y = objPos.y
                aDroid(psDroid).action = DACTION_MOVETODROIDREPAIR
                moveDroidTo psDroid, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y
            Else
                aDroid(psDroid).action = DACTION_NONE
            End If
        Else
            If Not droidUpdateDroidRepair(psDroid) Then
                droidRepairStopped castDroid(aDroid(psDroid).psActionTarget(0)), psDroid
                aDroid(psDroid).action = DACTION_NONE
                moveStopDroid psDroid
                'if the order is RTR then resubmit order so that the unit will go to repair facility point
                If orderState(psDroid, DORDER_RTR) Then
                    orderDroid psDroid, DORDER_RTR, ModeImmediate
                End If
            Else
                ' don't let the target for a repair shuffle
                If aDroid(aDroid(psDroid).psActionTarget(0)).sMove.Status = MOVESHUFFLE Then
                    moveStopDroid aDroid(psDroid).psActionTarget(0)
                End If
            End If
        End If
    Case DACTION_WAITFORREARM
        ' wait here for the rearm pad to instruct the vtol to move
        If aDroid(psDroid).psActionTarget(0) = 0 Then
            ' rearm pad destroyed - move to another
            objTrace aDroid(psDroid).id, "actionUpdateDroid", "rearm pad gone - switch to new one"
            moveToRearm psDroid
            GoTo switchEnd
        End If
        If DROID_STOPPED(psDroid) Then
            If vtolHappy(psDroid) Then
                objTrace aDroid(psDroid).id, "actionUpdateDroid", "do not need to rearm after all"
                ' don't actually need to rearm so just sit next to the rearm pad
                aDroid(psDroid).action = DACTION_NONE
            End If
        End If
    Case DACTION_CLEARREARMPAD
        If DROID_STOPPED(psDroid) Then
            aDroid(psDroid).action = DACTION_NONE
            objTrace aDroid(psDroid).id, "actionUpdateDroid", "clearing rearm pad"
            If Not vtolHappy(psDroid) Then  ' Droid has cleared the rearm pad without getting rearmed. One way this can happen if a rearming pad was built under the VTOL while it was waiting for a pad.
                moveToRearm psDroid  ' Rearm somewhere else instead.
            End If
        End If
    Case DACTION_WAITDURINGREARM
        ' this gets cleared by the rearm pad
    Case DACTION_MOVETOREARM
        If aDroid(psDroid).psActionTarget(0) = 0 Then
            ' base destroyed - find another
            objTrace aDroid(psDroid).id, "actionUpdateDroid", "rearm gone - find another"
            moveToRearm psDroid
            GoTo switchEnd
        End If

        If visibleObject(psDroid, aDroid(psDroid).psActionTarget(0), False) <> 0 Then
            psStruct = findNearestReArmPad(psDroid, aDroid(psDroid).psActionTarget(0), True)
            ' got close to the rearm pad - now find a clear one
            objTrace aDroid(psDroid).id, "actionUpdateDroid", "Seen rearm pad - searching for available one"

            If psStruct <> 0 Then
                ' found a clear landing pad - go for it
                objTrace aDroid(psDroid).id, "actionUpdateDroid", "Found clear rearm pad"
                setDroidActionTarget psDroid, psStruct, 0
            End If

            aDroid(psDroid).action = DACTION_WAITFORREARM
        End If

        If DROID_STOPPED(psDroid) Or aDroid(psDroid).action = DACTION_WAITFORREARM Then
            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
            rpos.x = objPos.x
            rpos.y = objPos.y
            If Not actionVTOLLandingPos(psDroid, rpos) Then
                ' totally bunged up - give up
                objTrace aDroid(psDroid).id, "actionUpdateDroid", "Couldn't find a clear tile near rearm pad - returning to base"
                orderDroid psDroid, DORDER_RTB, ModeImmediate
                GoTo switchEnd
            End If
            objTrace aDroid(psDroid).id, "actionUpdateDroid", "moving to rearm pad at %d,%d (%d,%d)", rpos.x, rpos.y, rpos.x \ TILE_UNITS, rpos.y \ TILE_UNITS
            moveDroidToDirect psDroid, rpos.x, rpos.y
        End If
    Case Else
        ASSERT_FAILURE "!""unknown action""", "src/action.cpp:2259", "actionUpdateDroid", "unknown action"
    End Select
switchEnd:

    If aDroid(psDroid).action <> DACTION_MOVEFIRE And _
       aDroid(psDroid).action <> DACTION_ATTACK And _
       aDroid(psDroid).action <> DACTION_MOVETOATTACK And _
       aDroid(psDroid).action <> DACTION_MOVETODROIDREPAIR And _
       aDroid(psDroid).action <> DACTION_DROIDREPAIR And _
       aDroid(psDroid).action <> DACTION_BUILD And _
       aDroid(psDroid).action <> DACTION_OBSERVE And _
       aDroid(psDroid).action <> DACTION_MOVETOOBSERVE Then
        'use 0 for all non-combat droid types
        If aDroid(psDroid).numWeaps = 0 Then
            If aDroid(psDroid).asWeaps(0).rot.direction <> 0 Or aDroid(psDroid).asWeaps(0).rot.pitch <> 0 Then
                actionAlignTurret psDroid, 0
            End If
        Else
            For i = 0 To aDroid(psDroid).numWeaps - 1
                If aDroid(psDroid).asWeaps(i).rot.direction <> 0 Or aDroid(psDroid).asWeaps(i).rot.pitch <> 0 Then
                    actionAlignTurret psDroid, i
                End If
            Next
        End If
    End If
    checkDroid psDroid, "src/action.cpp:2291", "actionUpdateDroid", max_check_object_recursion
End Sub

' (VB) actionUpdateDroid, DACTION_MOVEFIRE: bDirect && visGetBlockingWall(...) (only called when bDirect)
Private Function visGetBlockingWall_(ByVal bDirect As Boolean, ByVal psDroid As Long, ByVal psTarget As Long) As Boolean
    If bDirect Then
        visGetBlockingWall_ = visGetBlockingWall(psDroid, psTarget) <> 0
    End If
End Function

' (VB) actionUpdateDroid, DACTION_ATTACK, weapon i > 0: (order->type == DORDER_ATTACK || order->type == DORDER_ATTACKTARGET) &&
' psDroid->psActionTarget[i] != psDroid->psActionTarget[0] && validTarget(psDroid, psDroid->psActionTarget[0], i) &&
' actionInRange(psDroid, psDroid->psActionTarget[0], i), evaluated left to right with short-circuit
Private Function attackOrderCanShoot(ByVal psDroid As Long, ByVal i As Long) As Boolean
    If aDroid(psDroid).order.typ = DORDER_ATTACK Or aDroid(psDroid).order.typ = DORDER_ATTACKTARGET Then
        If aDroid(psDroid).psActionTarget(i) <> aDroid(psDroid).psActionTarget(0) Then
            If validTarget(psDroid, aDroid(psDroid).psActionTarget(0), i) Then
                attackOrderCanShoot = actionInRange(psDroid, aDroid(psDroid).psActionTarget(0), i)
            End If
        End If
    End If
End Function

' (VB) actionUpdateDroid, DACTION_ATTACK: actionVisibleTarget(psDroid, psActionTarget, i) && actionInRange(psDroid, psActionTarget, i)
Private Function actionVisibleTarget_(ByVal psDroid As Long, ByVal psActionTarget As Long, ByVal i As Long) As Boolean
    If actionVisibleTarget(psDroid, psActionTarget, i) Then
        actionVisibleTarget_ = actionInRange(psDroid, psActionTarget, i)
    End If
End Function

' (VB) actionUpdateDroid, DACTION_ATTACK: !psWeapStats->rotate || actionTargetTurret(psDroid, psActionTarget, &psDroid->asWeaps[i])
Private Function turretAlignedOrFixed(ByVal psDroid As Long, ByVal psActionTarget As Long, ByVal psWeapStats As Long, ByVal i As Long) As Boolean
    If asWeaponStats(psWeapStats).rotate = 0 Then
        turretAlignedOrFixed = True
    Else
        turretAlignedOrFixed = actionTargetTurret(psDroid, psActionTarget, i)
    End If
End Function

' (VB) actionUpdateDroid, DACTION_VTOLATTACK: vtolEmpty(psDroid) || psDroid->psActionTarget[0] == nullptr ||
' (electronicDroid(psDroid) && psDroid->player == psDroid->psActionTarget[0]->player) || !targetIsValid
Private Function vtolEmptyOrLost(ByVal psDroid As Long, ByVal targetIsValid As Boolean) As Boolean
    If vtolEmpty(psDroid) Then
        vtolEmptyOrLost = True
    ElseIf aDroid(psDroid).psActionTarget(0) = 0 Then
        vtolEmptyOrLost = True
    ElseIf electronicDroidSameOwner(psDroid) Then
        vtolEmptyOrLost = True
    Else
        ' Huh? !targetIsValid can't be true, we just checked for it
        vtolEmptyOrLost = Not targetIsValid
    End If
End Function

' (VB) electronicDroid(psDroid) && psDroid->player == psDroid->psActionTarget[i]->player
Private Function electronicDroidSameOwner(ByVal psDroid As Long, Optional ByVal i As Long = 0) As Boolean
    If electronicDroid(psDroid) Then
        electronicDroidSameOwner = aDroid(psDroid).player = SIMPLE_OBJECT_player(aDroid(psDroid).psActionTarget(i))
    End If
End Function

' Overall action function that is called by the specific action functions
Private Sub actionDroidBase(ByVal psDroid As Long, psAction As DROID_ACTION_DATA)
    Dim psWeapStats As Long, pos As Vector2i, secHoldActive As Boolean, actionIsInRange As Boolean
    Dim priorDroidAction As DROID_ACTION, priorActionTarget0 As Long, hasValidWeapon As Boolean, i As Long
    Dim pbx As Long, pby As Long, xdiff As Long, ydiff As Long, objPos As Vector3i, ok As Boolean

    If psAction.psObj <> 0 Then
        If Not (SIMPLE_OBJECT_died(psAction.psObj) = 0) Then ASSERT_FAILURE "psAction->psObj == nullptr || !psAction->psObj->died", "src/action.cpp:2297", "actionDroidBase", "Droid dead": Exit Sub
    End If

    psWeapStats = DROID_getWeaponStats(psDroid, 0)
    pos = Vector2i_(0, 0)

    checkDroid psDroid, "src/action.cpp:2302", "actionDroidBase", max_check_object_recursion

    secHoldActive = secondaryGetState(psDroid, DSO_HALTTYPE) = DSS_HALT_HOLD
    actionIsInRange = False
    aDroid(psDroid).actionStarted = gameTime
    syncDebugDroid_ "actionDroidBase", psDroid, Asc("-")
    syncDebug "actionDroidBase", "%d does %s", aDroid(psDroid).id, getDroidActionName(psAction.action)
    objTrace aDroid(psDroid).id, "actionDroidBase", "base set action to %s (was %s)", getDroidActionName(psAction.action), getDroidActionName(aDroid(psDroid).action)

    priorDroidAction = aDroid(psDroid).action
    priorActionTarget0 = aDroid(psDroid).psActionTarget(0)

    hasValidWeapon = False
    For i = 0 To MAX_WEAPONS - 1
        hasValidWeapon = hasValidWeapon Or validTarget(psDroid, psAction.psObj, i)
    Next
    Select Case psAction.action
    Case DACTION_NONE
        ' Clear up what ever the droid was doing before if necessary
        If Not DROID_STOPPED(psDroid) Then
            moveStopDroid psDroid
        End If
        aDroid(psDroid).action = DACTION_NONE
        aDroid(psDroid).actionPos = Vector2i_(0, 0)
        aDroid(psDroid).actionStarted = 0
        aDroid(psDroid).actionPoints = 0
        If aDroid(psDroid).numWeaps > 0 Then
            For i = 0 To aDroid(psDroid).numWeaps - 1
                setDroidActionTarget psDroid, 0, i
            Next
        Else
            setDroidActionTarget psDroid, 0, 0
        End If

    Case DACTION_TRANSPORTWAITTOFLYIN
        aDroid(psDroid).action = DACTION_TRANSPORTWAITTOFLYIN

    Case DACTION_ATTACK
        If aDroid(psDroid).asWeaps(0).nStat = 0 Or DROID_isTransporter(psDroid) Or psAction.psObj = psDroid Then
            GoTo switchEnd
        End If
        If Not hasValidWeapon Then
            ' continuing is pointless, we were given an invalid target
            ' for ex. AA gun can't attack ground unit
            GoTo switchEnd
        End If
        If vtolEmpty(psDroid) Then
            If DROID_isVtolRearming(psDroid) Then
                ' ignore attack order: we have no ammo and we are rearming right now, or going to rearm, or waiting for it
                GoTo switchEnd
            End If
        End If
        If electronicDroid(psDroid) Then
            'check for low or zero resistance - just zero resistance!
            If SIMPLE_OBJECT_type(psAction.psObj) = OBJ_STRUCTURE Then
                If Not validStructResistance(psAction.psObj) Then
                    'structure is low resistance already so don't attack
                    aDroid(psDroid).action = DACTION_NONE
                    GoTo switchEnd
                End If
            End If

            'in multiPlayer cannot electronically attack a transporter
            If bMultiPlayer Then
                If SIMPLE_OBJECT_type(psAction.psObj) = OBJ_DROID Then
                    If DROID_isTransporter(psAction.psObj) Then
                        aDroid(psDroid).action = DACTION_NONE
                        GoTo switchEnd
                    End If
                End If
            End If
        End If

        ' note the droid's current pos so that scout & patrol orders know how far the
        ' droid has gone during an attack
        ' slightly strange place to store this I know, but I didn't want to add any more to the droid
        aDroid(psDroid).actionPos.x = aDroid(psDroid).pos.x
        aDroid(psDroid).actionPos.y = aDroid(psDroid).pos.y
        setDroidActionTarget psDroid, psAction.psObj, 0
        actionIsInRange = actionInRange(psDroid, aDroid(psDroid).psActionTarget(0), 0, False)
        ok = False
        If secHoldActive Then
            If aDroid(psDroid).order.typ = DORDER_ATTACKTARGET _
               Or aDroid(psDroid).order.typ = DORDER_NONE _
               Or aDroid(psDroid).order.typ = DORDER_HOLD _
               Or aDroid(psDroid).order.typ = DORDER_FIRESUPPORT Then
                ok = True
            ElseIf aDroid(psDroid).order.typ = DORDER_GUARD Then
                ok = hasCommander(psDroid)
            End If
        End If
        If Not ok Then
            If Not DROID_isVtol(psDroid) Then
                ok = orderStateObj(psDroid, DORDER_FIRESUPPORT) <> 0
            End If
        End If
        If ok Then
            aDroid(psDroid).action = DACTION_ATTACK ' holding, try attack straightaway
        ElseIf actionInsideMinRange(psDroid, psAction.psObj, psWeapStats) Then ' too close?
            If Not proj_Direct(psWeapStats) Then
                If asWeaponStats(psWeapStats).rotate <> 0 Then
                    aDroid(psDroid).action = DACTION_ATTACK
                Else
                    aDroid(psDroid).action = DACTION_ROTATETOATTACK
                    objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
                    moveTurnDroid psDroid, objPos.x, objPos.y
                End If
            ElseIf aDroid(psDroid).order.typ <> DORDER_HOLD And secondaryGetState(psDroid, DSO_HALTTYPE) <> DSS_HALT_HOLD Then
                pbx = 0
                pby = 0
                ' direct fire - try and extend the range
                aDroid(psDroid).action = DACTION_MOVETOATTACK
                actionCalcPullBackPoint psDroid, psAction.psObj, pbx, pby

                turnOffMultiMsg True
                moveDroidTo psDroid, pbx, pby
                turnOffMultiMsg False
            End If
        ElseIf aDroid(psDroid).order.typ <> DORDER_HOLD _
               And secondaryGetState(psDroid, DSO_HALTTYPE) <> DSS_HALT_HOLD _
               And Not actionIsInRange Then ' approach closer?
            aDroid(psDroid).action = DACTION_MOVETOATTACK
            turnOffMultiMsg True
            objPos = SIMPLE_OBJECT_pos(psAction.psObj)
            moveDroidTo psDroid, objPos.x, objPos.y
            turnOffMultiMsg False
        ElseIf actionIsInRange Or (aDroid(psDroid).order.typ <> DORDER_HOLD And secondaryGetState(psDroid, DSO_HALTTYPE) = DSS_HALT_HOLD) Then
            If Not DROID_isVtol(psDroid) Then
                moveStopDroid psDroid
                aDroid(psDroid).action = DACTION_ATTACK
            Else
                ' this will make droid lift up before attacking
                ' and also make them go to rearm pad if low ammo
                aDroid(psDroid).action = DACTION_MOVETOATTACK
                turnOffMultiMsg True
                objPos = SIMPLE_OBJECT_pos(psAction.psObj)
                moveDroidTo psDroid, objPos.x, objPos.y
                turnOffMultiMsg False
            End If
        End If

    Case DACTION_MOVETOREARM
        aDroid(psDroid).action = DACTION_MOVETOREARM
        objPos = SIMPLE_OBJECT_pos(psAction.psObj)
        aDroid(psDroid).actionPos.x = objPos.x
        aDroid(psDroid).actionPos.y = objPos.y
        aDroid(psDroid).actionStarted = gameTime
        setDroidActionTarget psDroid, psAction.psObj, 0
        objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
        pos.x = objPos.x
        pos.y = objPos.y
        If Not actionVTOLLandingPos(psDroid, pos) Then
            ' totally bunged up - give up
            objTrace aDroid(psDroid).id, "actionDroidBase", "move to rearm action failed!"
            orderDroid psDroid, DORDER_RTB, ModeImmediate
            GoTo switchEnd
        End If
        objTrace aDroid(psDroid).id, "actionDroidBase", "move to rearm"
        moveDroidToDirect psDroid, pos.x, pos.y
    Case DACTION_CLEARREARMPAD
        debug_ 2474, LOG_NEVER, "actionDroidBase", "Unit %d clearing rearm pad", aDroid(psDroid).id
        aDroid(psDroid).action = DACTION_CLEARREARMPAD
        setDroidActionTarget psDroid, psAction.psObj, 0
        objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
        pos.x = objPos.x
        pos.y = objPos.y
        If Not actionVTOLLandingPos(psDroid, pos) Then
            ' totally bunged up - give up
            objTrace aDroid(psDroid).id, "actionDroidBase", "clear rearm pad action failed!"
            orderDroid psDroid, DORDER_RTB, ModeImmediate
            GoTo switchEnd
        End If
        objTrace aDroid(psDroid).id, "actionDroidBase", "move to clear rearm pad"
        moveDroidToDirect psDroid, pos.x, pos.y
    Case DACTION_MOVE, DACTION_TRANSPORTIN, DACTION_TRANSPORTOUT, DACTION_RETURNTOPOS, DACTION_FIRESUPPORT_RETREAT
        aDroid(psDroid).action = psAction.action
        aDroid(psDroid).actionPos.x = psAction.x
        aDroid(psDroid).actionPos.y = psAction.y
        aDroid(psDroid).actionStarted = gameTime
        setDroidActionTarget psDroid, psAction.psObj, 0
        moveDroidTo psDroid, psAction.x, psAction.y

    Case DACTION_BUILD
        If aDroid(psDroid).order.psStats = -1 Then
            aDroid(psDroid).action = DACTION_NONE
            GoTo switchEnd
        End If
        'ASSERT_OR_RETURN(, order->type == DORDER_BUILD || order->type == DORDER_HELPBUILD || order->type == DORDER_LINEBUILD, "cannot start build action without a build order");
        If Not (psAction.x <> 0 And psAction.y <> 0) Then ASSERT_FAILURE "psAction->x > 0 && psAction->y > 0", "src/action.cpp:2508", "actionDroidBase", "Bad build order position": Exit Sub
        aDroid(psDroid).action = DACTION_MOVETOBUILD
        aDroid(psDroid).actionPos.x = psAction.x
        aDroid(psDroid).actionPos.y = psAction.y
        ' Move a droid in a tile if it has never been moved before
        If aDroid(psDroid).actionPos.x = aDroid(psDroid).pos.x And aDroid(psDroid).actionPos.y = aDroid(psDroid).pos.y Then
            actionRemoveDroidsFromBuildPos aDroid(psDroid).player, aDroid(psDroid).actionPos, aDroid(psDroid).order.direction, structureStatsBase(aDroid(psDroid).order.psStats)
        Else
            moveDroidToNoFormation psDroid, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y
        End If
    Case DACTION_DEMOLISH
        If Not (aDroid(psDroid).order.typ = DORDER_DEMOLISH) Then ASSERT_FAILURE "order->type == DORDER_DEMOLISH", "src/action.cpp:2523", "actionDroidBase", "cannot start demolish action without a demolish order": Exit Sub
        aDroid(psDroid).action = DACTION_MOVETODEMOLISH
        aDroid(psDroid).actionPos.x = psAction.x
        aDroid(psDroid).actionPos.y = psAction.y
        ok = aDroid(psDroid).order.psObj <> 0
        If ok Then ok = SIMPLE_OBJECT_type(aDroid(psDroid).order.psObj) = OBJ_STRUCTURE
        If Not ok Then ASSERT_FAILURE "(order->psObj != nullptr) && (order->psObj->type == OBJ_STRUCTURE)", "src/action.cpp:2527", "actionDroidBase", "invalid target for demolish order": Exit Sub
        aDroid(psDroid).order.psStats = aStruct(aDroid(psDroid).order.psObj).pStructureType
        setDroidActionTarget psDroid, psAction.psObj, 0
        moveDroidTo psDroid, psAction.x, psAction.y
    Case DACTION_REPAIR
        aDroid(psDroid).action = psAction.action
        aDroid(psDroid).actionPos.x = psAction.x
        aDroid(psDroid).actionPos.y = psAction.y
        'this needs setting so that automatic repair works
        setDroidActionTarget psDroid, psAction.psObj, 0
        ok = aDroid(psDroid).psActionTarget(0) <> 0
        If ok Then ok = SIMPLE_OBJECT_type(aDroid(psDroid).psActionTarget(0)) = OBJ_STRUCTURE
        If Not ok Then ASSERT_FAILURE "(psDroid->psActionTarget[0] != nullptr) && (psDroid->psActionTarget[0]->type == OBJ_STRUCTURE)", "src/action.cpp:2538", "actionDroidBase", "invalid target for repair order": Exit Sub
        aDroid(psDroid).order.psStats = aStruct(aDroid(psDroid).psActionTarget(0)).pStructureType
        If secHoldActive And (aDroid(psDroid).order.typ = DORDER_NONE Or aDroid(psDroid).order.typ = DORDER_HOLD) Then
            aDroid(psDroid).action = DACTION_REPAIR
        ElseIf (Not secHoldActive And aDroid(psDroid).order.typ <> DORDER_HOLD) Or (secHoldActive And aDroid(psDroid).order.typ = DORDER_REPAIR) Then
            aDroid(psDroid).action = DACTION_MOVETOREPAIR
            moveDroidTo psDroid, psAction.x, psAction.y
        End If
    Case DACTION_OBSERVE
        aDroid(psDroid).action = psAction.action
        setDroidActionTarget psDroid, psAction.psObj, 0
        aDroid(psDroid).actionPos.x = aDroid(psDroid).pos.x
        aDroid(psDroid).actionPos.y = aDroid(psDroid).pos.y
        If secondaryGetState(psDroid, DSO_HALTTYPE) <> DSS_HALT_GUARD And (aDroid(psDroid).order.typ = DORDER_NONE Or aDroid(psDroid).order.typ = DORDER_HOLD) Then
            If visibleObject(psDroid, aDroid(psDroid).psActionTarget(0), False) <> 0 Then
                aDroid(psDroid).action = DACTION_OBSERVE
            Else
                aDroid(psDroid).action = DACTION_NONE
            End If
        ElseIf Not cbSensorDroid(psDroid) And ((Not secHoldActive And aDroid(psDroid).order.typ <> DORDER_HOLD) Or (secHoldActive And aDroid(psDroid).order.typ = DORDER_OBSERVE)) Then
            aDroid(psDroid).action = DACTION_MOVETOOBSERVE
            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).psActionTarget(0))
            moveDroidTo psDroid, objPos.x, objPos.y
        End If
    Case DACTION_FIRESUPPORT
        aDroid(psDroid).action = DACTION_FIRESUPPORT
        If Not DROID_isVtol(psDroid) And Not secHoldActive Then
            If SIMPLE_OBJECT_type(aDroid(psDroid).order.psObj) <> OBJ_STRUCTURE Then
                objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj)
                moveDroidTo psDroid, objPos.x, objPos.y  ' movetotarget.
            End If
        End If
    Case DACTION_SULK
        aDroid(psDroid).action = DACTION_SULK
        ' hmmm, hope this doesn't cause any problems!
        aDroid(psDroid).actionStarted = gameTime + MIN_SULK_TIME + (gameRand(MAX_SULK_TIME - MIN_SULK_TIME))
    Case DACTION_WAITFORREPAIR
        aDroid(psDroid).action = DACTION_WAITFORREPAIR
        ' set the time so we can tell whether the start the self repair or not
        aDroid(psDroid).actionStarted = gameTime
    Case DACTION_MOVETOREPAIRPOINT
        aDroid(psDroid).action = psAction.action
        aDroid(psDroid).actionPos.x = psAction.x
        aDroid(psDroid).actionPos.y = psAction.y
        aDroid(psDroid).actionStarted = gameTime
        setDroidActionTarget psDroid, psAction.psObj, 0
        moveDroidToNoFormation psDroid, psAction.x, psAction.y
    Case DACTION_WAITDURINGREPAIR
        aDroid(psDroid).action = DACTION_WAITDURINGREPAIR
    Case DACTION_MOVETOREARMPOINT
        objTrace aDroid(psDroid).id, "actionDroidBase", "set to move to rearm pad"
        aDroid(psDroid).action = psAction.action
        aDroid(psDroid).actionPos.x = psAction.x
        aDroid(psDroid).actionPos.y = psAction.y
        aDroid(psDroid).actionStarted = gameTime
        setDroidActionTarget psDroid, psAction.psObj, 0
        moveDroidToDirect psDroid, psAction.x, psAction.y

        ' make sure there aren't any other VTOLs on the rearm pad
        ensureRearmPadClear psAction.psObj, psDroid
    Case DACTION_DROIDREPAIR
        aDroid(psDroid).action = psAction.action
        aDroid(psDroid).actionPos.x = psAction.x
        aDroid(psDroid).actionPos.y = psAction.y
        setDroidActionTarget psDroid, psAction.psObj, 0
        'initialise the action points
        aDroid(psDroid).actionPoints = 0
        aDroid(psDroid).actionStarted = gameTime
        xdiff = aDroid(psDroid).pos.x - psAction.x
        ydiff = aDroid(psDroid).pos.y - psAction.y
        If secHoldActive And (aDroid(psDroid).order.typ = DORDER_NONE Or aDroid(psDroid).order.typ = DORDER_HOLD) Then
            aDroid(psDroid).action = DACTION_DROIDREPAIR
        ElseIf ((Not secHoldActive And aDroid(psDroid).order.typ <> DORDER_HOLD) Or (secHoldActive And aDroid(psDroid).order.typ = DORDER_DROIDREPAIR)) _
               And ((xdiff * xdiff + ydiff * ydiff) > REPAIR_RANGE * REPAIR_RANGE) Then  ' check that we actually need to move closer
            aDroid(psDroid).action = DACTION_MOVETODROIDREPAIR
            moveDroidTo psDroid, psAction.x, psAction.y
        End If
        If aDroid(psDroid).action = DACTION_DROIDREPAIR And priorDroidAction <> DACTION_DROIDREPAIR Then
            droidRepairStarted castDroid(aDroid(psDroid).psActionTarget(0)), psDroid
        End If
    Case DACTION_RESTORE
        If Not (aDroid(psDroid).order.typ = DORDER_RESTORE) Then ASSERT_FAILURE "order->type == DORDER_RESTORE", "src/action.cpp:2635", "actionDroidBase", "cannot start restore action without a restore order": Exit Sub
        aDroid(psDroid).action = psAction.action
        aDroid(psDroid).actionPos.x = psAction.x
        aDroid(psDroid).actionPos.y = psAction.y
        ok = aDroid(psDroid).order.psObj <> 0
        If ok Then ok = SIMPLE_OBJECT_type(aDroid(psDroid).order.psObj) = OBJ_STRUCTURE
        If Not ok Then ASSERT_FAILURE "(order->psObj != nullptr) && (order->psObj->type == OBJ_STRUCTURE)", "src/action.cpp:2639", "actionDroidBase", "invalid target for restore order": Exit Sub
        aDroid(psDroid).order.psStats = aStruct(aDroid(psDroid).order.psObj).pStructureType
        setDroidActionTarget psDroid, psAction.psObj, 0
        If aDroid(psDroid).order.typ <> DORDER_HOLD Then
            aDroid(psDroid).action = DACTION_MOVETORESTORE
            moveDroidTo psDroid, psAction.x, psAction.y
        End If
    Case Else
        ASSERT_FAILURE "!""unknown action""", "src/action.cpp:2649", "actionDroidBase", "actionUnitBase: unknown action"
    End Select
switchEnd:
    If priorDroidAction = DACTION_DROIDREPAIR And aDroid(psDroid).action <> DACTION_DROIDREPAIR Then
        droidRepairStopped castDroid(priorActionTarget0), psDroid
    End If
    syncDebugDroid_ "actionDroidBase", psDroid, Asc("+")
    checkDroid psDroid, "src/action.cpp:2657", "actionDroidBase", max_check_object_recursion
End Sub

' Give a droid an action
Public Sub actionDroid(ByVal psDroid As Long, ByVal action As DROID_ACTION)
    Dim sAction As DROID_ACTION_DATA

    sAction.action = action
    actionDroidBase psDroid, sAction
End Sub

' Give a droid an action with a location target
Public Sub actionDroid_loc(ByVal psDroid As Long, ByVal action As DROID_ACTION, ByVal x As Long, ByVal y As Long)
    Dim sAction As DROID_ACTION_DATA

    sAction.action = action
    sAction.x = x
    sAction.y = y
    actionDroidBase psDroid, sAction
End Sub

' Give a droid an action with an object target
Public Sub actionDroid_obj(ByVal psDroid As Long, ByVal action As DROID_ACTION, ByVal psObj As Long)
    Dim sAction As DROID_ACTION_DATA, objPos As Vector3i

    sAction.action = action
    sAction.psObj = psObj
    objPos = SIMPLE_OBJECT_pos(psObj)
    sAction.x = objPos.x
    sAction.y = objPos.y
    actionDroidBase psDroid, sAction
End Sub

' Give a droid an action with an object target and a location
Public Sub actionDroid_objLoc(ByVal psDroid As Long, ByVal action As DROID_ACTION, _
                              ByVal psObj As Long, ByVal x As Long, ByVal y As Long)
    Dim sAction As DROID_ACTION_DATA

    sAction.action = action
    sAction.psObj = psObj
    sAction.x = x
    sAction.y = y
    actionDroidBase psDroid, sAction
End Sub

' send the vtol droid back to the nearest rearming pad - if one otherwise
' return to base
Public Sub moveToRearm(ByVal psDroid As Long)
    Dim psStruct As Long

    checkDroid psDroid, "src/action.cpp:2715", "moveToRearm", max_check_object_recursion

    If Not DROID_isVtol(psDroid) Then
        Exit Sub
    End If

    'if droid is already returning - ignore
    If DROID_isVtolRearming(psDroid) Then
        Exit Sub
    End If

    'get the droid to fly back to a ReArming Pad
    ' don't worry about finding a clear one for the minute
    psStruct = findNearestReArmPad(psDroid, aDroid(psDroid).psBaseStruct, False)
    If psStruct <> 0 Then
        ' note a base rearm pad if the vtol doesn't have one
        If aDroid(psDroid).psBaseStruct = 0 Then
            setDroidBase psDroid, psStruct
        End If

        'return to re-arming pad
        If aDroid(psDroid).order.typ = DORDER_NONE Then
            ' no order set - use the rearm order to ensure the unit goes back
            ' to the landing pad
            orderDroidObj psDroid, DORDER_REARM, psStruct, ModeImmediate
        Else
            actionDroid_obj psDroid, DACTION_MOVETOREARM, psStruct
        End If
    Else
        'return to base un-armed
        objTrace aDroid(psDroid).id, "moveToRearm", "Did not find an available rearm pad - RTB instead"
        orderDroid psDroid, DORDER_RTB, ModeImmediate
    End If
End Sub

' whether a tile is suitable for a vtol to land on
Private Function vtolLandingTile(ByVal x As Long, ByVal y As Long) As Boolean
    Dim psTile As Long
    If x < 0 Or x >= mapWidth Or y < 0 Or y >= mapHeight Then
        vtolLandingTile = False
        Exit Function
    End If

    psTile = mapTile(x, y)
    If (psMapTiles(psTile).tileInfoBits And BITS_FPATHBLOCK) <> 0 Or _
       TileIsOccupied(psTile) Or _
       terrainType(psTile) = TER_CLIFFFACE Or _
       terrainType(psTile) = TER_WATER Then
        vtolLandingTile = False
        Exit Function
    End If
    vtolLandingTile = True
End Function

' (VB) match(x, y, matchState)
Private Function tileMatch(ByVal match As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Select Case match
    Case TILE_MATCH_VTOL_LANDING
        tileMatch = vtolLandingTileSearchFunction(x, y)
    End Select
End Function

' Performs a space-filling spiral-like search from startX,startY up to (and
' including) radius. For each tile, the search function is called; if it
' returns 'true', the search will finish immediately.
' returns true if finished because the searchFunction requested termination,
' false if the radius limit was reached
Private Function spiralSearch(ByVal startX As Long, ByVal startY As Long, ByVal max_radius As Long, ByVal match As Long) As Boolean
    Dim radius As Long, min_distance As Long, max_distance As Long, dx As Long, dy As Long, distance As Long

    ' test center tile
    If tileMatch(match, startX, startY) Then
        spiralSearch = True
        Exit Function
    End If

    ' test for each radius, from 1 to max_radius (inclusive)
    For radius = 1 To max_radius
        ' choose tiles that are between radius and radius+1 away from center
        ' distances are squared
        min_distance = radius * radius
        max_distance = min_distance + 2 * radius

        ' dx starts with 1, to visiting tiles on same row or col as start twice
        For dx = 1 To max_radius
            For dy = 0 To max_radius
                ' Current distance, squared
                distance = dx * dx + dy * dy

                ' Ignore tiles outside of the current circle
                If distance < min_distance Or distance > max_distance Then
                    GoTo continue_
                End If

                ' call search function for each of the 4 quadrants of the circle
                If tileMatch(match, startX + dx, startY + dy) Then spiralSearch = True: Exit Function
                If tileMatch(match, startX - dx, startY - dy) Then spiralSearch = True: Exit Function
                If tileMatch(match, startX + dy, startY - dx) Then spiralSearch = True: Exit Function
                If tileMatch(match, startX - dy, startY + dx) Then spiralSearch = True: Exit Function
continue_:
            Next
        Next
    Next

    spiralSearch = False
End Function

' an internal tileMatchFunction that checks if x and y are coordinates of a
' valid landing place. matchState: a Vector2i where these coordintates should be stored
' returns true if coordinates are a valid landing tile, false if not.
Private Function vtolLandingTileSearchFunction(ByVal x As Long, ByVal y As Long) As Boolean
    If vtolLandingTile(x, y) Then
        landingXY.x = x
        landingXY.y = y
        vtolLandingTileSearchFunction = True
        Exit Function
    End If

    vtolLandingTileSearchFunction = False
End Function

' Choose a landing position for a VTOL when it goes to rearm that is close to rearm
' pad but not on it, since it may be busy by the time we get there.
Public Function actionVTOLLandingPos(ByVal psDroid As Long, p As Vector2i) As Boolean
    Dim startX As Long, startY As Long, it As Long, psCurr As Long, t As Vector2i, foundTile As Boolean

    checkDroid psDroid, "src/action.cpp:2870", "actionVTOLLandingPos", max_check_object_recursion

    ' Initial box dimensions and set iteration count to zero
    startX = map_coord(p.x)
    startY = map_coord(p.y)

    ' set blocking flags for all the other droids
    it = apsDroidLists(aDroid(psDroid).player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If DROID_STOPPED(psCurr) Then
            t.x = map_coord(aDroid(psCurr).pos.x)
            t.y = map_coord(aDroid(psCurr).pos.y)
        Else
            t.x = map_coord(aDroid(psCurr).sMove.destination.x)
            t.y = map_coord(aDroid(psCurr).sMove.destination.y)
        End If
        If psCurr <> psDroid Then
            If tileOnMap(t.x, t.y) Then
                psMapTiles(mapTile(t.x, t.y)).tileInfoBits = psMapTiles(mapTile(t.x, t.y)).tileInfoBits Or BITS_FPATHBLOCK
            End If
        End If
        it = lnNext(it)
    Loop

    ' search for landing tile; will stop when found or radius exceeded
    landingXY.x = 0
    landingXY.y = 0
    foundTile = spiralSearch(startX, startY, vtolLandingRadius, TILE_MATCH_VTOL_LANDING)
    If foundTile Then
        objTrace aDroid(psDroid).id, "actionVTOLLandingPos", "Unit %d landing pos (%d,%d)", aDroid(psDroid).id, landingXY.x, landingXY.y
        p.x = world_coord(landingXY.x) + TILE_UNITS \ 2
        p.y = world_coord(landingXY.y) + TILE_UNITS \ 2
    End If

    ' clear blocking flags for all the other droids
    it = apsDroidLists(aDroid(psDroid).player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If DROID_STOPPED(psCurr) Then
            t.x = map_coord(aDroid(psCurr).pos.x)
            t.y = map_coord(aDroid(psCurr).pos.y)
        Else
            t.x = map_coord(aDroid(psCurr).sMove.destination.x)
            t.y = map_coord(aDroid(psCurr).sMove.destination.y)
        End If
        If tileOnMap(t.x, t.y) Then
            psMapTiles(mapTile(t.x, t.y)).tileInfoBits = psMapTiles(mapTile(t.x, t.y)).tileInfoBits And Not BITS_FPATHBLOCK
        End If
        it = lnNext(it)
    Loop

    actionVTOLLandingPos = foundTile
End Function
