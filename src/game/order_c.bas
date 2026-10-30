Attribute VB_Name = "order_c"
' src/order.cpp, order.h: Functions for setting the orders of a droid or
' group of droids.
' (VB) RtrBestResult is a UDT with the constructor functions
' RtrBestResult_ and RtrBestResult_order. findSomeoneToRepair's
' std::priority_queue with its RSComparator only ever gives its top, the
' greatest droid in the comparator's order (which is total, the ids
' breaking ties): a linear search for it (RSComparator_less). Overloads:
' findSomeoneToRepair(const STRUCTURE *, int) keeps the name,
' findSomeoneToRepair(const DROID *, int) is findSomeoneToRepair_droid;
' highestQueuedModule(DroidOrder const &, ...) is highestQueuedModule_order.
' checkForDamagedStruct's std::pair<STRUCTURE *, DROID_ACTION> is its
' result plus a ByRef action. The DEBUG-only secondaryPrintFactories is
' not ported (a release build).
Option Explicit

' How long a droid runs after it fails do respond due to low moral.
Private Const RUN_TIME As Long = 8000

' How long a droid runs burning after it fails do respond due to low moral.
Private Const RUN_BURN_TIME As Long = 10000

' The distance a droid has in guard mode.
Private Const DEFEND_MAXDIST As Long = TILE_UNITS * 3

' The distance a droid has in guard mode.
' @todo seems to be used as equivalent to GUARD_MAXDIST.
Private Const DEFEND_BASEDIST As Long = TILE_UNITS * 3

' The distance a droid has in guard mode. Equivalent to GUARD_MAXDIST, but used for droids being on a command group.
Private Const DEFEND_CMD_MAXDIST As Long = TILE_UNITS * 8

' The distance a droid has in guard mode. Equivalent to GUARD_BASEDIST, but used for droids being on a command group.
Private Const DEFEND_CMD_BASEDIST As Long = TILE_UNITS * 5

' The maximum distance a constructor droid has in guard mode.
Private Const CONSTRUCT_MAXDIST As Long = TILE_UNITS * 8

' The maximum distance allowed to a droid to move out of the path on a patrol/scout.
Private Const SCOUT_DIST As Long = TILE_UNITS * 8

' The maximum distance allowed to a droid to move out of the path if already attacking a target on a patrol/scout.
Private Const SCOUT_ATTACK_DIST As Long = TILE_UNITS * 5

' The minimum delay to be used on orderPlayFireSupportAudio() for fire support sound.
Private Const AUDIO_DELAY_FIRESUPPORT As Long = 3 * GAME_TICKS_PER_SEC

' the center of this area starts at the closest repair droid/facility!
Private Const MAGIC_SUITABLE_REPAIR_AREA As Long = (REPAIR_RANGE * 5) * (REPAIR_RANGE * 5)

' Whether an order effect has been displayed
' @todo better documentation required.
Private bOrderEffectDisplayed As Boolean

Private Type RtrBestResult
    typ As RTR_DATA_TYPE
    psObj As Long                       ' BASE_OBJECT *
End Type

' RSComparator's state
Private Type RSComparator
    selfPosX As Long
    selfPosY As Long
    selfPlayer As Long
End Type

' decideWhereToRepairAndBalance's static vectors (to save allocations)
Private vFacilityPos() As Vector3i
Private vFacility() As Long
Private vFacilityCloseEnough() As Long
Private vDroidPos() As Vector3i
Private vDroid() As Long
Private vDroidCloseEnough() As Long
Private vFacilityPos_size As Long, vFacility_size As Long, vFacilityCloseEnough_size As Long
Private vDroidPos_size As Long, vDroid_size As Long, vDroidCloseEnough_size As Long

Private gridListRepair() As Long        ' _findSomeoneToRepair's GridList
Private gridListDamaged() As Long       ' checkForDamagedStruct's GridList

' RtrBestResult(RTR_DATA_TYPE type, BASE_OBJECT *psObj)
Private Function RtrBestResult_(ByVal typ As RTR_DATA_TYPE, ByVal psObj As Long) As RtrBestResult
    RtrBestResult_.typ = typ
    RtrBestResult_.psObj = psObj
End Function

' RtrBestResult(DROID_ORDER_DATA *psOrder)
Private Function RtrBestResult_order(psOrder As DROID_ORDER_DATA) As RtrBestResult
    RtrBestResult_order.typ = psOrder.rtrType
    RtrBestResult_order.psObj = psOrder.psObj
    If SIMPLE_OBJECT_type(RtrBestResult_order.psObj) = OBJ_STRUCTURE Then
        If asStructureStats(aStruct(RtrBestResult_order.psObj).pStructureType).typ = REF_REPAIR_FACILITY Then
            RtrBestResult_order.typ = RTR_TYPE_REPAIR_FACILITY
        Else
            RtrBestResult_order.typ = RTR_TYPE_HQ
        End If
    Else
        RtrBestResult_order.typ = RTR_TYPE_DROID
    End If
End Function

' This function checks if the droid is off range. If yes, it uses actionDroid() to make the droid to move to its target if its target is on range, or to move to its order position if not.
' @todo droid doesn't shoot while returning to the guard position.
Private Sub orderCheckGuardPosition(ByVal psDroid As Long, ByVal range As Long)
    Dim x As Long, y As Long, xdiff As Long, ydiff As Long, useMove As Boolean, objPos As Vector3i

    If aDroid(psDroid).order.psObj <> 0 Then
        ' repair droids always follow behind - don't want them jumping into the line of fire
        useMove = False
        If (Not (aDroid(psDroid).droidType = DROID_REPAIR Or aDroid(psDroid).droidType = DROID_CYBORG_REPAIR)) Then
            If SIMPLE_OBJECT_type(aDroid(psDroid).order.psObj) = OBJ_DROID Then
                useMove = orderStateLoc(aDroid(psDroid).order.psObj, DORDER_MOVE, x, y)
            End If
        End If
        If useMove Then
            ' got a moving droid - check against where the unit is going
            aDroid(psDroid).order.pos = Vector2i_(x, y)
        Else
            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj)
            aDroid(psDroid).order.pos = Vector2i_(objPos.x, objPos.y)
        End If
    End If

    xdiff = aDroid(psDroid).pos.x - aDroid(psDroid).order.pos.x
    ydiff = aDroid(psDroid).pos.y - aDroid(psDroid).order.pos.y
    If xdiff * xdiff + ydiff * ydiff > range * range Then
        If (aDroid(psDroid).sMove.Status <> MOVEINACTIVE) And _
           ((aDroid(psDroid).action = DACTION_MOVE) Or _
            (aDroid(psDroid).action = DACTION_MOVEFIRE)) Then
            xdiff = aDroid(psDroid).sMove.destination.x - aDroid(psDroid).order.pos.x
            ydiff = aDroid(psDroid).sMove.destination.y - aDroid(psDroid).order.pos.y
            If xdiff * xdiff + ydiff * ydiff > range * range Then
                actionDroid_loc psDroid, DACTION_MOVE, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
            End If
        Else
            actionDroid_loc psDroid, DACTION_MOVE, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
        End If
    End If
End Sub

' RSComparator(const STRUCTURE *self), RSComparator(const DROID *self)
Private Function RSComparator_(ByVal self As Long) As RSComparator
    Dim pos As Vector3i
    pos = SIMPLE_OBJECT_pos(self)
    RSComparator_.selfPosX = pos.x
    RSComparator_.selfPosY = pos.y
    RSComparator_.selfPlayer = SIMPLE_OBJECT_player(self)
End Function

' RSComparator's distanceCheck lambda
Private Function RSComparator_distanceCheck(cmp As RSComparator, ByVal l As Long, ByVal r As Long) As Boolean
    Dim ldist As Long, rdist As Long
    ldist = (aDroid(l).pos.x - cmp.selfPosX) * (aDroid(l).pos.x - cmp.selfPosX) + (aDroid(l).pos.y - cmp.selfPosY) * (aDroid(l).pos.y - cmp.selfPosY)
    rdist = (aDroid(r).pos.x - cmp.selfPosX) * (aDroid(r).pos.x - cmp.selfPosX) + (aDroid(r).pos.y - cmp.selfPosY) * (aDroid(r).pos.y - cmp.selfPosY)
    ' debug(LOG_REPAIRS, "comparator called %i %i:  (ldist %i >= %i rdist) %i", l->id, r->id, ldist, rdist, ldist >= rdist);
    If ldist <> rdist Then
        RSComparator_distanceCheck = ldist > rdist
        Exit Function
    End If
    ' If the distances are the same, fallback to damage level check: prefer the one which is more damaged.
    If aDroid(l).body <> aDroid(r).body Then
        RSComparator_distanceCheck = ULT(aDroid(r).body, aDroid(l).body)
        Exit Function
    End If
    ' Fall back to comparing droid IDs, which should never be the same, as the last resort
    ' to preserve the consistency of comparator (e.g. strict weak ordering).
    RSComparator_distanceCheck = ULT(aDroid(r).id, aDroid(l).id)
End Function

' "less" comparator for our priority queue
Private Function RSComparator_less(cmp As RSComparator, ByVal l As Long, ByVal r As Long) As Boolean
    Dim lown As Boolean, rown As Boolean, lTarget As Long, rTarget As Long, lhighest As Boolean, rhighest As Boolean
    Dim lsecond As Boolean, rsecond As Boolean

    ' Here, we know that both left and right Droids are:
    ' - either our own, or one of our allies
    ' - already within repair radius
    ' - not dead
    ' - don't have full HP

    ' always prefer our own units
    lown = aDroid(l).player = cmp.selfPlayer
    rown = aDroid(r).player = cmp.selfPlayer
    If lown And Not rown Then RSComparator_less = True: Exit Function
    If Not lown And rown Then RSComparator_less = False: Exit Function

    ' Next highest priority:
    ' Take any droid with orders to Return to Repair (DORDER_RTR),
    ' or that have been ordered to a specific repair facility (DORDER_RTR_SPECIFIED)
    lTarget = orderStateObj(l, DORDER_RTR)
    rTarget = orderStateObj(r, DORDER_RTR)
    lhighest = ((aDroid(l).order.typ = DORDER_RTR Or aDroid(l).order.typ = DORDER_RTR_SPECIFIED))
    rhighest = ((aDroid(r).order.typ = DORDER_RTR Or aDroid(r).order.typ = DORDER_RTR_SPECIFIED))

    If lhighest And Not rhighest Then RSComparator_less = True: Exit Function
    If Not lhighest And rhighest Then RSComparator_less = False: Exit Function

    If lhighest And rhighest Then
        ' break the tie with distance check
        RSComparator_less = RSComparator_distanceCheck(cmp, l, r)
        Exit Function
    End If

    ' Second highest priority:
    ' Help out another nearby repair facility
    lsecond = waitsForRepairFacility(l, lTarget)
    rsecond = waitsForRepairFacility(r, rTarget)
    If lsecond And Not rsecond Then RSComparator_less = True: Exit Function
    If Not lsecond And rsecond Then RSComparator_less = False: Exit Function

    ' at this point we don't really have a preference
    ' just repair closest.
    ' In case they have the same distance,
    ' fallback to the damage level check: prefer the most damaged.
    RSComparator_less = RSComparator_distanceCheck(cmp, l, r)
End Function

' (VB) RSComparator: l->action == DACTION_WAITFORREPAIR && lTarget && lTarget->type == OBJ_STRUCTURE &&
' ((STRUCTURE *)lTarget)->pStructureType && ((STRUCTURE *)lTarget)->pStructureType->type == REF_REPAIR_FACILITY
Private Function waitsForRepairFacility(ByVal l As Long, ByVal lTarget As Long) As Boolean
    If aDroid(l).action = DACTION_WAITFORREPAIR And lTarget <> 0 Then
        If SIMPLE_OBJECT_type(lTarget) = OBJ_STRUCTURE Then
            If aStruct(lTarget).pStructureType <> -1 Then
                waitsForRepairFacility = asStructureStats(aStruct(lTarget).pStructureType).typ = REF_REPAIR_FACILITY
            End If
        End If
    End If
End Function

' @return return top priority droid to repair, or nullptr
' psRepairFac: the repair facility's STRUCTURE (0 = nullptr), whose pFunctionality->repairFacility C++ passes
Private Function findSomeoneToRepair_(ByVal psRepairFac As Long, cmp As RSComparator, _
                                      ByVal x As Long, ByVal y As Long, ByVal radius As Long, ByVal player As Long) As Long
    Dim n As Long, gi As Long, psDroid As Long, top As Long
    n = gridStartIterateRepairCandidates(x, y, radius, player, gridListRepair)
    top = 0
    For gi = 0 To n - 1
        psDroid = gridListRepair(gi)
        If DROID_isDamaged(psDroid) Then
            ' queue.push(psDroid)
            If top = 0 Then
                top = psDroid
            ElseIf RSComparator_less(cmp, top, psDroid) Then
                top = psDroid
            End If
        ElseIf aDroid(psDroid).order.typ = DORDER_RTR Or aDroid(psDroid).order.typ = DORDER_RTR_SPECIFIED Then
            ' we intentionally iterate over fully healthy droids too, in order to clear RTR order
            aDroid(psDroid).body = MIN_(aDroid(psDroid).originalBody, aDroid(psDroid).body)
            droidWasFullyRepaired psDroid, psRepairFac
        End If
    Next
    If top <> 0 Then
        debug_ 272, LOG_REPAIRS, "_findSomeoneToRepair", "top was %i", aDroid(top).id
        findSomeoneToRepair_ = top
        Exit Function
    End If
    findSomeoneToRepair_ = 0
End Function

Public Function findSomeoneToRepair(ByVal psStructure As Long, ByVal radius As Long) As Long
    Dim cmpr As RSComparator
    cmpr = RSComparator_(psStructure)
    findSomeoneToRepair = findSomeoneToRepair_(psStructure, cmpr, aStruct(psStructure).pos.x, aStruct(psStructure).pos.y, radius, aStruct(psStructure).player)
End Function

Public Function findSomeoneToRepair_droid(ByVal psTurret As Long, ByVal radius As Long) As Long
    Dim cmpr As RSComparator
    cmpr = RSComparator_(psTurret)
    findSomeoneToRepair_droid = findSomeoneToRepair_(0, cmpr, aDroid(psTurret).pos.x, aDroid(psTurret).pos.y, radius, aDroid(psTurret).player)
End Function

' This function checks if there are any structures to repair or help build in a given radius near the droid defined by REPAIR_RANGE if it is on hold, and REPAIR_MAXDIST if not on hold.
' It returns a damaged or incomplete structure if any was found or nullptr if none was found.
Private Function checkForDamagedStruct(ByVal psDroid As Long, bestAction As DROID_ACTION) As Long
    Dim psFailedTarget As Long, radius As Long, bestDistanceSq As Long, best As Long
    Dim n As Long, gi As Long, object As Long, distanceSq As Long, structure_ As Long

    psFailedTarget = 0
    If aDroid(psDroid).action = DACTION_SULK Then
        psFailedTarget = aDroid(psDroid).psActionTarget(0)
    End If

    If (aDroid(psDroid).order.typ = DORDER_HOLD) Or (aDroid(psDroid).order.typ = DORDER_NONE And secondaryGetState(psDroid, DSO_HALTTYPE) = DSS_HALT_HOLD) Then
        radius = REPAIR_RANGE
    Else
        radius = REPAIR_MAXDIST
    End If

    bestDistanceSq = radius * radius
    best = 0
    bestAction = DACTION_NONE

    n = gridStartIterate(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, radius, gridListDamaged)
    For gi = 0 To n - 1
        object = gridListDamaged(gi)
        distanceSq = droidSqDist(psDroid, object)  ' droidSqDist returns -1 if unreachable, (unsigned)-1 is a big number.

        structure_ = castStructure(object)
        If structure_ = 0 Then GoTo continue_  ' Must be a structure.
        If structure_ = psFailedTarget Then GoTo continue_  ' Must not have just failed to reach it.
        If ULT(bestDistanceSq, distanceSq) Then GoTo continue_  ' Must be as close as possible.
        If visibleObject(psDroid, structure_, False) = 0 Then GoTo continue_  ' Must be able to sense it.
        If Not aiCheckAlliances(aDroid(psDroid).player, aStruct(structure_).player) Then GoTo continue_  ' Must be a friendly structure.
        If checkDroidsDemolishing(structure_) Then GoTo continue_  ' Must not be trying to get rid of it.

        ' Check for structures to repair.
        If aStruct(structure_).status = SS_BUILT And STRUCTURE_isDamaged(structure_) Then
            bestDistanceSq = distanceSq
            best = structure_
            bestAction = DACTION_REPAIR
        ' Check for structures to help build.
        ElseIf aStruct(structure_).status = SS_BEING_BUILT Then
            bestDistanceSq = distanceSq
            best = structure_
            bestAction = DACTION_BUILD
        End If
continue_:
    Next

    checkForDamagedStruct = best
End Function

Private Function isRepairlikeAction(ByVal action As DROID_ACTION) As Boolean
    Select Case action
    Case DACTION_BUILD, DACTION_BUILDWANDER, DACTION_DEMOLISH, DACTION_DROIDREPAIR, DACTION_MOVETOBUILD, _
         DACTION_MOVETODEMOLISH, DACTION_MOVETODROIDREPAIR, DACTION_MOVETOREPAIR, DACTION_MOVETORESTORE, _
         DACTION_REPAIR, DACTION_RESTORE
        isRepairlikeAction = True
    Case Else
        isRepairlikeAction = False
    End Select
End Function

Private Function tryDoRepairlikeAction(ByVal psDroid As Long) As Boolean
    Dim radius As Long, repairTarget As Long, damaged As Long, damagedAction As DROID_ACTION

    If isRepairlikeAction(aDroid(psDroid).action) Then
        tryDoRepairlikeAction = True  ' Already doing something.
        Exit Function
    End If
    Select Case aDroid(psDroid).droidType
    Case DROID_REPAIR, DROID_CYBORG_REPAIR
        If ((aDroid(psDroid).order.typ = DORDER_HOLD) Or _
            (aDroid(psDroid).order.typ = DORDER_NONE And secondaryGetState(psDroid, DSO_HALTTYPE) = DSS_HALT_HOLD)) Then
            radius = REPAIR_RANGE
        Else
            radius = REPAIR_MAXDIST
        End If
        'repair droids default to repairing droids within a given range
        repairTarget = findSomeoneToRepair_droid(psDroid, radius)
        If repairTarget <> 0 Then
            actionDroid_obj psDroid, DACTION_DROIDREPAIR, repairTarget
        End If
    Case DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT
        'construct droids default to repairing and helping structures within a given range
        damaged = checkForDamagedStruct(psDroid, damagedAction)
        If damagedAction = DACTION_REPAIR Then
            actionDroid_obj psDroid, damagedAction, damaged
        ElseIf damagedAction = DACTION_BUILD Then
            aDroid(psDroid).order.psStats = aStruct(damaged).pStructureType
            aDroid(psDroid).order.direction = aStruct(damaged).rot.direction
            actionDroid_loc psDroid, damagedAction, aStruct(damaged).pos.x, aStruct(damaged).pos.y
        End If
    Case Else
        tryDoRepairlikeAction = False
        Exit Function
    End Select
    tryDoRepairlikeAction = True
End Function

' This function updates all the orders status, according with psdroid's current order and state.
' Returns false if all further droid processing should be shortcut for this frame (because, for example, the droid is a transporter that was moved to the off-world list)
Public Function orderUpdateDroid(ByVal psDroid As Long) As Boolean
    Dim psObj As Long, psStruct As Long, psWall As Long, xdiff As Long, ydiff As Long, bAttack As Boolean
    Dim xoffset As Long, yoffset As Long, psWeapStats As Long, tooFarFromPath As Boolean, delta As Vector2i
    Dim attack As Boolean, tmp As Vector2i, edgeDiff As Vector2i, diff As Vector2i, angle As Long
    Dim temp As Long, transporter As Long, lb As LineBuild, psFireTarget As Long, psSpotter As Long
    Dim c As Boolean, objPos As Vector3i, commander As Long

    If psDroid = 0 Then
        orderUpdateDroid = True
        Exit Function
    ElseIf isDead(psDroid) Then
        orderUpdateDroid = True
        Exit Function
    End If

    psWeapStats = DROID_getWeaponStats(psDroid, 0)

    ' clear the target if it has died
    If aDroid(psDroid).order.psObj <> 0 Then
        If SIMPLE_OBJECT_died(aDroid(psDroid).order.psObj) <> 0 Then
            syncDebugObject_ "orderUpdateDroid", aDroid(psDroid).order.psObj, Asc("-")
            setDroidTarget psDroid, 0
            objTrace aDroid(psDroid).id, "orderUpdateDroid", "Target dead"
        End If
    End If

    'clear its base struct if its died
    If aDroid(psDroid).psBaseStruct <> 0 Then
        If SIMPLE_OBJECT_died(aDroid(psDroid).psBaseStruct) <> 0 Then
            syncDebugStructure_ "orderUpdateDroid", aDroid(psDroid).psBaseStruct, Asc("-")
            setDroidBase psDroid, 0
            objTrace aDroid(psDroid).id, "orderUpdateDroid", "Base struct dead"
        End If
    End If

    ' check for died objects in the list
    orderCheckList psDroid

    Select Case aDroid(psDroid).order.typ
    Case DORDER_NONE, DORDER_HOLD
        ' see if there are any orders queued up
        If orderDroidList(psDroid) Then
            ' started a new order, quit
            GoTo switchEnd
        ' if you are in a command group, default to guarding the commander
        ElseIf hasCommander(psDroid) And aDroid(psDroid).order.typ <> DORDER_HOLD _
               And aDroid(psDroid).order.psStats <> structGetDemolishStat() Then  ' stop the constructor auto repairing when it is about to demolish
            orderDroidObj psDroid, DORDER_GUARD, aGroup(aDroid(psDroid).psGroup).psCommander, ModeImmediate
        ElseIf DROID_isTransporter(psDroid) And Not bMultiPlayer Then

        ' default to guarding
        ElseIf Not tryDoRepairlikeAction(psDroid) Then
            If aDroid(psDroid).order.typ <> DORDER_HOLD _
               And aDroid(psDroid).order.psStats <> structGetDemolishStat() _
               And secondaryGetState(psDroid, DSO_HALTTYPE) = DSS_HALT_GUARD _
               And Not DROID_isVtol(psDroid) Then
                orderDroidLoc psDroid, DORDER_GUARD, aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, ModeImmediate
            End If
        End If
    Case DORDER_TRANSPORTRETURN
        If aDroid(psDroid).action = DACTION_NONE Then
            missionMoveTransporterOffWorld psDroid

            ' clear order
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)

            orderUpdateDroid = False ' signal to caller to skip further processing this frame - droid was moved to a different list!
            Exit Function
        End If
    Case DORDER_TRANSPORTOUT
        If aDroid(psDroid).action = DACTION_NONE Then
            If aDroid(psDroid).player = selectedPlayer Then
                If getDroidsToSafetyFlag() Then
                    'move droids in Transporter into holding list
                    moveDroidsToSafety psDroid
                    'we need the transporter to just sit off world for a while...
                    orderDroid psDroid, DORDER_TRANSPORTIN, ModeImmediate
                    ' set action transporter waits for timer
                    actionDroid psDroid, DACTION_TRANSPORTWAITTOFLYIN

                    missionSetReinforcementTime gameTime
                    'don't do this until waited for the required time
                    'fly Transporter back to get some more droids
                    'orderDroidLoc( psDroid, DORDER_TRANSPORTIN,
                    'getLandingX(selectedPlayer), getLandingY(selectedPlayer));
                Else
                    ' clear order
                    aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
                End If
                'the script can call startMission for this callback for offworld missions (if we want to change level)
                triggerEvent TRIGGER_TRANSPORTER_EXIT, psDroid

                aDroid(psDroid).sMove.speed = 0 ' Prevent radical movement vector when adjusting from home to away map exit and entry coordinates.
            End If
        End If
    Case DORDER_TRANSPORTIN
        If (aDroid(psDroid).action = DACTION_NONE) And _
           (aDroid(psDroid).sMove.Status = MOVEINACTIVE) Then
            ' clear order
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)

'FFS! You only wan't to do this if the droid being tracked IS the transporter! Not all the time!
' What about if your happily playing the game and tracking a droid, and re-enforcements come in!
' And suddenly BLAM!!!! It drops you out of camera mode for no apparent reason! TOTALY CONFUSING
' THE PLAYER!
'
' Just had to get that off my chest....end of rant.....
'
            If psDroid = getTrackingDroid() Then    ' Thats better...
                ' deselect transporter if have been tracking
                If getWarCamStatus() Then
                    camToggleStatus
                End If
            End If

            DeSelectDroid psDroid

            'don't try the unload if moving droids to safety and still got some
            'droids left  - wait until full and then launch again
            c = False
            If aDroid(psDroid).player = selectedPlayer Then
                If getDroidsToSafetyFlag() Then
                    c = missionDroidsRemaining(selectedPlayer)
                End If
            End If
            If c Then
                resetTransporter
                triggerEvent TRIGGER_TRANSPORTER_LANDED, psDroid
            Else
                unloadTransporter psDroid, aDroid(psDroid).pos.x, aDroid(psDroid).pos.y
            End If
        End If
    Case DORDER_MOVE
        ' Just wait for the action to finish then clear the order
        If aDroid(psDroid).action = DACTION_NONE Or aDroid(psDroid).action = DACTION_ATTACK Then
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
        End If
    Case DORDER_RECOVER
        If aDroid(psDroid).order.psObj = 0 Then
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
        ElseIf aDroid(psDroid).action = DACTION_NONE Then
            ' stopped moving, but still haven't got the artifact
            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj)
            actionDroid_loc psDroid, DACTION_MOVE, objPos.x, objPos.y
        End If
    Case DORDER_SCOUT, DORDER_PATROL
        ' if there is an enemy around, attack it
        If aDroid(psDroid).action = DACTION_MOVE Or aDroid(psDroid).action = DACTION_MOVEFIRE Or (aDroid(psDroid).action = DACTION_NONE And DROID_isVtol(psDroid)) Then
            tooFarFromPath = False
            If DROID_isVtol(psDroid) And aDroid(psDroid).order.typ = DORDER_PATROL Then
                ' Don't stray too far from the patrol path - only attack if we're near it
                ' A fun algorithm to detect if we're near the path
                delta.x = aDroid(psDroid).order.pos.x - aDroid(psDroid).order.pos2.x
                delta.y = aDroid(psDroid).order.pos.y - aDroid(psDroid).order.pos2.y
                If delta.x = 0 And delta.y = 0 Then
                    tooFarFromPath = False
                ElseIf Abs(delta.x) >= Abs(delta.y) And _
                       MIN_(aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos2.x) - SCOUT_DIST <= aDroid(psDroid).pos.x And _
                       aDroid(psDroid).pos.x <= MAX_(aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos2.x) + SCOUT_DIST Then
                    tooFarFromPath = (Abs((aDroid(psDroid).pos.x - aDroid(psDroid).order.pos.x) * delta.y \ delta.x + _
                                          aDroid(psDroid).order.pos.y - aDroid(psDroid).pos.y) > SCOUT_DIST)
                ElseIf Abs(delta.x) <= Abs(delta.y) And _
                       MIN_(aDroid(psDroid).order.pos.y, aDroid(psDroid).order.pos2.y) - SCOUT_DIST <= aDroid(psDroid).pos.y And _
                       aDroid(psDroid).pos.y <= MAX_(aDroid(psDroid).order.pos.y, aDroid(psDroid).order.pos2.y) + SCOUT_DIST Then
                    tooFarFromPath = (Abs((aDroid(psDroid).pos.y - aDroid(psDroid).order.pos.y) * delta.x \ delta.y + _
                                          aDroid(psDroid).order.pos.x - aDroid(psDroid).pos.x) > SCOUT_DIST)
                Else
                    tooFarFromPath = True
                End If
            End If
            If Not tooFarFromPath Then
                ' true if in condition to set actionDroid to attack/observe
                attack = False
                If secondaryGetState(psDroid, DSO_ATTACK_LEVEL) = DSS_ALEV_ALWAYS Then
                    attack = aiBestNearestTarget(psDroid, psObj, 0, SCOUT_ATTACK_DIST) >= 0
                End If
                Select Case aDroid(psDroid).droidType
                Case DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT, DROID_REPAIR, DROID_CYBORG_REPAIR
                    tryDoRepairlikeAction psDroid
                Case DROID_WEAPON, DROID_CYBORG, DROID_CYBORG_SUPER, DROID_PERSON, DROID_COMMAND
                    If attack Then
                        actionDroid_obj psDroid, DACTION_ATTACK, psObj
                    End If
                Case DROID_SENSOR
                    If Not cbSensorDroid(psDroid) And attack Then
                        actionDroid_obj psDroid, DACTION_OBSERVE, psObj
                    End If
                Case Else
                    actionDroid psDroid, DACTION_NONE
                End Select
            End If
        End If
        If aDroid(psDroid).action = DACTION_NONE Then
            xdiff = aDroid(psDroid).pos.x - aDroid(psDroid).order.pos.x
            ydiff = aDroid(psDroid).pos.y - aDroid(psDroid).order.pos.y
            If xdiff * xdiff + ydiff * ydiff < SCOUT_DIST * SCOUT_DIST Then
                If aDroid(psDroid).order.typ = DORDER_PATROL Then
                    ' see if we have anything queued up
                    If orderDroidList(psDroid) Then
                        ' started a new order, quit
                        GoTo switchEnd
                    End If
                    If DROID_isVtol(psDroid) Then
                        If Not vtolFull(psDroid) And (aDroid(psDroid).secondaryOrder And DSS_ALEV_MASK) <> DSS_ALEV_NEVER Then
                            moveToRearm psDroid
                            GoTo switchEnd
                        End If
                    End If
                    ' head back to the other point
                    tmp = aDroid(psDroid).order.pos
                    aDroid(psDroid).order.pos = aDroid(psDroid).order.pos2
                    aDroid(psDroid).order.pos2 = tmp
                    actionDroid_loc psDroid, DACTION_MOVE, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
                Else
                    aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
                End If
            Else
                actionDroid_loc psDroid, DACTION_MOVE, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
            End If
        ElseIf ((aDroid(psDroid).action = DACTION_ATTACK) Or _
                (aDroid(psDroid).action = DACTION_VTOLATTACK) Or _
                (aDroid(psDroid).action = DACTION_MOVETOATTACK) Or _
                (aDroid(psDroid).action = DACTION_ROTATETOATTACK) Or _
                (aDroid(psDroid).action = DACTION_OBSERVE) Or _
                (aDroid(psDroid).action = DACTION_MOVETOOBSERVE)) And _
               secondaryGetState(psDroid, DSO_HALTTYPE) <> DSS_HALT_PURSUE Then
            ' attacking something - see if the droid has gone too far, go up to twice the distance we want to go, so that we don't repeatedly turn back when the target is almost in range.
            If objPosDiffSq(aDroid(psDroid).pos, Vector3i_(aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y, 0)) > (SCOUT_ATTACK_DIST * 2 * SCOUT_ATTACK_DIST * 2) Then
                actionDroid_loc psDroid, DACTION_RETURNTOPOS, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y
            End If
        End If
        If aDroid(psDroid).order.typ = DORDER_PATROL And DROID_isVtol(psDroid) Then
            If vtolEmpty(psDroid) And (aDroid(psDroid).secondaryOrder And DSS_ALEV_MASK) <> DSS_ALEV_NEVER Then
                moveToRearm psDroid  ' Completely empty (and we're not set to hold fire), don't bother patrolling.
                GoTo switchEnd
            End If
        End If
    Case DORDER_CIRCLE
        ' if there is an enemy around, attack it
        c = False
        If aDroid(psDroid).action = DACTION_MOVE Then
            If secondaryGetState(psDroid, DSO_ATTACK_LEVEL) = DSS_ALEV_ALWAYS Then
                c = aiBestNearestTarget(psDroid, psObj, 0, SCOUT_ATTACK_DIST) >= 0
            End If
        End If
        If c Then
            Select Case aDroid(psDroid).droidType
            Case DROID_WEAPON, DROID_CYBORG, DROID_CYBORG_SUPER, DROID_PERSON, DROID_COMMAND
                actionDroid_obj psDroid, DACTION_ATTACK, psObj
            Case DROID_SENSOR
                actionDroid_obj psDroid, DACTION_OBSERVE, psObj
            Case Else
                actionDroid psDroid, DACTION_NONE
            End Select
        ElseIf aDroid(psDroid).action = DACTION_NONE Or aDroid(psDroid).action = DACTION_MOVE Then
            If aDroid(psDroid).action = DACTION_MOVE Then
                ' see if we have anything queued up
                If orderDroidList(psDroid) Then
                    ' started a new order, quit
                    GoTo switchEnd
                End If
            End If

            edgeDiff.x = aDroid(psDroid).pos.x - aDroid(psDroid).actionPos.x
            edgeDiff.y = aDroid(psDroid).pos.y - aDroid(psDroid).actionPos.y
            If aDroid(psDroid).action <> DACTION_MOVE Or dot2i(edgeDiff, edgeDiff) <= TILE_UNITS * 4 * TILE_UNITS * 4 Then
                'Watermelon:use orderX,orderY as local space origin and calculate droid direction in local space
                diff.x = aDroid(psDroid).pos.x - aDroid(psDroid).order.pos.x
                diff.y = aDroid(psDroid).pos.y - aDroid(psDroid).order.pos.y
                angle = (iAtan2(diff.x, diff.y) - DEG(30)) And &HFFFF&
                Do
                    xoffset = iSinR(angle, 1500)
                    yoffset = iCosR(angle, 1500)
                    angle = (angle - DEG(10)) And &HFFFF&
                Loop While Not worldOnMap(aDroid(psDroid).order.pos.x + xoffset, aDroid(psDroid).order.pos.y + yoffset)    ' Don't try to fly off map.
                actionDroid_loc psDroid, DACTION_MOVE, aDroid(psDroid).order.pos.x + xoffset, aDroid(psDroid).order.pos.y + yoffset
            End If

            If DROID_isVtol(psDroid) Then
                If vtolEmpty(psDroid) And (aDroid(psDroid).secondaryOrder And DSS_ALEV_MASK) <> DSS_ALEV_NEVER Then
                    moveToRearm psDroid  ' Completely empty (and we're not set to hold fire), don't bother circling.
                    GoTo switchEnd
                End If
            End If
        ElseIf ((aDroid(psDroid).action = DACTION_ATTACK) Or _
                (aDroid(psDroid).action = DACTION_VTOLATTACK) Or _
                (aDroid(psDroid).action = DACTION_MOVETOATTACK) Or _
                (aDroid(psDroid).action = DACTION_ROTATETOATTACK) Or _
                (aDroid(psDroid).action = DACTION_OBSERVE) Or _
                (aDroid(psDroid).action = DACTION_MOVETOOBSERVE)) And _
               secondaryGetState(psDroid, DSO_HALTTYPE) <> DSS_HALT_PURSUE Then
            ' attacking something - see if the droid has gone too far
            xdiff = aDroid(psDroid).pos.x - aDroid(psDroid).actionPos.x
            ydiff = aDroid(psDroid).pos.y - aDroid(psDroid).actionPos.y
            If xdiff * xdiff + ydiff * ydiff > 2000 * 2000 Then
                ' head back to the target location
                actionDroid_loc psDroid, DACTION_RETURNTOPOS, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
            End If
        End If
    Case DORDER_HELPBUILD, DORDER_DEMOLISH, DORDER_OBSERVE, DORDER_REPAIR, DORDER_DROIDREPAIR, DORDER_RESTORE
        If aDroid(psDroid).action = DACTION_NONE Or aDroid(psDroid).order.psObj = 0 Then
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            actionDroid psDroid, DACTION_NONE
            If aDroid(psDroid).player = selectedPlayer Then
                intRefreshScreen
            End If
        End If
    Case DORDER_REARM
        If aDroid(psDroid).order.psObj = 0 Or aDroid(psDroid).psActionTarget(0) = 0 Then
            ' arm pad destroyed find another
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            moveToRearm psDroid
        ElseIf aDroid(psDroid).action = DACTION_NONE Then
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
        End If
    Case DORDER_ATTACK, DORDER_ATTACKTARGET
        c = aDroid(psDroid).order.psObj = 0
        If Not c Then c = SIMPLE_OBJECT_died(aDroid(psDroid).order.psObj) <> 0
        If c Then
            ' if vtol then return to rearm pad as long as there are no other
            ' orders queued up
            If DROID_isVtol(psDroid) Then
                If Not orderDroidList(psDroid) Then
                    aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
                    moveToRearm psDroid
                End If
            Else
                aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
                actionDroid psDroid, DACTION_NONE
            End If
            GoTo switchEnd
        End If
        c = False
        If (aDroid(psDroid).action = DACTION_MOVE) Or _
           (aDroid(psDroid).action = DACTION_MOVEFIRE) Then
            If actionVisibleTarget(psDroid, aDroid(psDroid).order.psObj, 0) Then c = Not DROID_isVtol(psDroid)
        End If
        If c Then
            ' moved near enough to attack change to attack action
            actionDroid_obj psDroid, DACTION_ATTACK, aDroid(psDroid).order.psObj
            GoTo switchEnd
        End If
        c = False
        If (aDroid(psDroid).action = DACTION_MOVETOATTACK) And _
           Not DROID_isVtol(psDroid) Then
            If Not actionVisibleTarget(psDroid, aDroid(psDroid).order.psObj, 0) Then
                c = secondaryGetState(psDroid, DSO_HALTTYPE) <> DSS_HALT_HOLD
            End If
        End If
        If c Then
            ' lost sight of the target while chasing it - change to a move action so
            ' that the unit will fire on other things while moving
            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj)
            actionDroid_loc psDroid, DACTION_MOVE, objPos.x, objPos.y
            GoTo switchEnd
        End If
        c = False
        psWall = 0
        If Not DROID_isVtol(psDroid) _
           And aDroid(psDroid).order.psObj = aDroid(psDroid).psActionTarget(0) Then
            If actionInRange(psDroid, aDroid(psDroid).order.psObj, 0) Then
                psWall = visGetBlockingWall(psDroid, aDroid(psDroid).order.psObj)
                If psWall <> 0 Then
                    If Not aiCheckAlliances(aStruct(psWall).player, aDroid(psDroid).player) Then
                        ' this is still wrong and should be unified with "action.cpp: actionUpdateDroid: case DACTION_ATTACK",
                        ' because "case DACTION_ATTACK" actualy takes care of checking for all weapons (not only the first one)
                        ' and for direct/indirect weapons.
                        If psWeapStats <> -1 Then c = proj_Direct(psWeapStats)
                    End If
                End If
            End If
        End If
        If c Then
            ' there is a wall in the way - attack that
            actionDroid_obj psDroid, DACTION_ATTACK, psWall
        ElseIf (aDroid(psDroid).action = DACTION_NONE) Or _
               (aDroid(psDroid).action = DACTION_CLEARREARMPAD) Then
            c = False
            If (aDroid(psDroid).order.typ = DORDER_ATTACKTARGET Or aDroid(psDroid).order.typ = DORDER_ATTACK) _
               And secondaryGetState(psDroid, DSO_HALTTYPE) = DSS_HALT_HOLD Then
                c = Not actionInRange(psDroid, aDroid(psDroid).order.psObj, 0)
            End If
            If c Then
                ' target is not in range and DSS_HALT_HOLD: give up, don't move
                aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            ElseIf Not DROID_isVtol(psDroid) Then
                actionDroid_obj psDroid, DACTION_ATTACK, aDroid(psDroid).order.psObj
            ElseIf allVtolsRearmed(psDroid) Then
                actionDroid_obj psDroid, DACTION_ATTACK, aDroid(psDroid).order.psObj
            End If
        End If
    Case DORDER_BUILD
        If aDroid(psDroid).action = DACTION_BUILD And _
           aDroid(psDroid).order.psObj = 0 Then
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            actionDroid psDroid, DACTION_NONE
            objTrace aDroid(psDroid).id, "orderUpdateDroid", "Clearing build order since build target is gone"
        ElseIf aDroid(psDroid).action = DACTION_NONE Then
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            objTrace aDroid(psDroid).id, "orderUpdateDroid", "Clearing build order since build action is reset"
        End If
    Case DORDER_EMBARK
        ' only place it can be trapped - in multiPlayer can only put cyborgs onto a Cyborg Transporter
        temp = aDroid(psDroid).order.psObj ' NOTE: It is possible to have a NULL here

        c = False
        If temp <> 0 Then
            If aDroid(temp).droidType = DROID_TRANSPORTER Then c = Not DROID_isCyborg(psDroid)
        End If
        If c Then
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            actionDroid psDroid, DACTION_NONE
            If aDroid(psDroid).player = selectedPlayer Then
                audio_PlayBuildFailedOnce
                addConsoleMessage "We can't do that! We must be a Cyborg unit to use a Cyborg Transport!", DEFAULT_JUSTIFY, selectedPlayer
            End If
        Else
            ' don't want the droids to go into a formation for this order
            If aDroid(psDroid).sMove.psFormation <> 0 Then
                formationLeave aDroid(psDroid).sMove.psFormation, psDroid
                aDroid(psDroid).sMove.psFormation = 0
            End If

            ' Wait for the action to finish then assign to Transporter (if not already flying)
            c = aDroid(psDroid).order.psObj = 0
            If Not c Then c = transporterFlying(aDroid(psDroid).order.psObj)
            If c Then
                aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
                actionDroid psDroid, DACTION_NONE
            ElseIf Abs(aDroid(psDroid).pos.x - SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj).x) < TILE_UNITS _
                   And Abs(aDroid(psDroid).pos.y - SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj).y) < TILE_UNITS Then
                ' save the target of current droid (the transporter)
                transporter = aDroid(psDroid).order.psObj

                ' Make sure that it really is a valid droid
                checkDroid transporter, "src/order.cpp:920", "orderUpdateDroid", max_check_object_recursion

                ' order the droid to stop so moveUpdateDroid does not process this unit
                orderDroid psDroid, DORDER_STOP, ModeImmediate
                setDroidTarget psDroid, 0
                aDroid(psDroid).order.psObj = 0
                secondarySetState psDroid, DSO_RETURN_TO_LOC, DSS_NONE
                moveReallyStopDroid psDroid

                ' Fire off embark event
                triggerEvent TRIGGER_TRANSPORTER_EMBARKED, transporter

                ' We must add the droid to the transporter only *after*
                ' processing changing its orders (see above).
                transporterAddDroid transporter, psDroid
            ElseIf aDroid(psDroid).action = DACTION_NONE Then
                objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj)
                actionDroid_loc psDroid, DACTION_MOVE, objPos.x, objPos.y
            End If
        End If
        ' Do we need to clear the secondary order "DSO_EMBARK" here? (PD)
    Case DORDER_DISEMBARK
        'only valid in multiPlayer mode
        If bMultiPlayer Then
            'this order can only be given to Transporter droids
            If DROID_isTransporter(psDroid) Then
                ' once the Transporter has reached its destination (and landed),
                ' get all the units to disembark
                If aDroid(psDroid).action <> DACTION_MOVE And aDroid(psDroid).action <> DACTION_MOVEFIRE And _
                   aDroid(psDroid).sMove.Status = MOVEINACTIVE And aDroid(psDroid).sMove.iVertSpeed = 0 Then
                    unloadTransporter psDroid, aDroid(psDroid).pos.x, aDroid(psDroid).pos.y
                    'reset the transporter's order
                    aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
                End If
            End If
        End If
    Case DORDER_RTB
        ' Just wait for the action to finish then clear the order
        If aDroid(psDroid).action = DACTION_NONE Then
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            secondarySetState psDroid, DSO_RETURN_TO_LOC, DSS_NONE
        End If
    Case DORDER_RTR, DORDER_RTR_SPECIFIED
        ' send them back to commander, no need to repair
        If Not DROID_isDamaged(psDroid) Then
            objTrace aDroid(psDroid).id, "orderUpdateDroid", "was RTR, but we are full health"
            droidWasFullyRepaired psDroid, 0
        End If

        If aDroid(psDroid).order.psObj = 0 Then
            objTrace aDroid(psDroid).id, "orderUpdateDroid", "Our RTR target got lost. Let's try again."
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            orderDroid psDroid, DORDER_RTR, ModeImmediate
        ElseIf aDroid(psDroid).action = DACTION_NONE Then
            ' get repair facility pointer
            psStruct = aDroid(psDroid).order.psObj
            If Not (psStruct <> 0) Then ASSERT_FAILURE "psStruct != nullptr", "src/order.cpp:991", "orderUpdateDroid", "orderUpdateUnit: invalid structure pointer"

            If objPosDiffSq(aDroid(psDroid).pos, SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj)) < REPAIR_RANGE * REPAIR_RANGE Then
                ' action droid to wait
                actionDroid psDroid, DACTION_WAITFORREPAIR
            Else
                ' move the droid closer to the repair point
                ' setting target to null will trigger search for nearest repair point: we might have a better option after all
                aDroid(psDroid).order.psObj = 0
            End If
        End If
        If (aDroid(psDroid).action = DACTION_WAITFORREPAIR Or aDroid(psDroid).action = DACTION_WAITDURINGREPAIR) And _
           aDroid(psDroid).order.psObj <> 0 Then
            If objPosDiffSq(aDroid(psDroid).pos, SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj)) > REPAIR_RANGE * REPAIR_RANGE Then
                ' was being repaired, but somehow got lost. recalculate reparing point
                aDroid(psDroid).order.psObj = 0
            End If
        End If
    Case DORDER_LINEBUILD
        If aDroid(psDroid).action = DACTION_NONE Or _
           (aDroid(psDroid).action = DACTION_BUILD And aDroid(psDroid).order.psObj = 0) Then
            ' finished building the current structure
            lb = calcLineBuild(aDroid(psDroid).order.psStats, aDroid(psDroid).order.direction, aDroid(psDroid).order.pos, aDroid(psDroid).order.pos2)
            If lb.count <= 1 Then
                ' finished all the structures - done
                aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
                GoTo switchEnd
            End If

            ' update the position for another structure
            aDroid(psDroid).order.pos = LineBuild_at(lb, 1)

            ' build another structure
            setDroidTarget psDroid, 0
            actionDroid_loc psDroid, DACTION_BUILD, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
            'intRefreshScreen();
        End If
    Case DORDER_FIRESUPPORT
        If aDroid(psDroid).order.psObj = 0 Then
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            If DROID_isVtol(psDroid) Then
                ' VTOLs need special care. We don't want them landing in the
                ' middle of the battlefield if the sensor died!!! Rearm or RTB.
                moveToRearm psDroid
            Else
                actionDroid psDroid, DACTION_NONE
            End If
            GoTo switchEnd
        End If
        'before targetting - check if VTOL's are empty or need to retreat to repair
        c = False
        If DROID_isVtol(psDroid) Then
            c = vtolEmpty(psDroid)
            If Not c Then c = secondaryCheckDamageLevelDeselect(psDroid, secondaryGetState(psDroid, DSO_REPAIR_LEVEL))
        End If
        If c Then
            moveToRearm psDroid
        'indirect weapon droid attached to (standard)sensor droid
        Else
            psFireTarget = 0

            If SIMPLE_OBJECT_type(aDroid(psDroid).order.psObj) = OBJ_DROID Then
                psSpotter = aDroid(psDroid).order.psObj

                If aDroid(psSpotter).action = DACTION_OBSERVE _
                   Or (aDroid(psSpotter).droidType = DROID_COMMAND And aDroid(psSpotter).action = DACTION_ATTACK) Then
                    psFireTarget = aDroid(psSpotter).psActionTarget(0)
                End If
            ElseIf SIMPLE_OBJECT_type(aDroid(psDroid).order.psObj) = OBJ_STRUCTURE Then
                psSpotter = aDroid(psDroid).order.psObj

                psFireTarget = aStruct(psSpotter).psTarget(0)
            End If

            c = False
            If psFireTarget <> 0 Then
                If SIMPLE_OBJECT_died(psFireTarget) = 0 Then c = checkAnyWeaponsTarget(psDroid, psFireTarget)
            End If
            If c Then
                bAttack = False
                If DROID_isVtol(psDroid) Then
                    If aDroid(psDroid).action = DACTION_WAITDURINGREARM Then
                        ' Calm down. It's not that serious.
                    ElseIf Not vtolEmpty(psDroid) And _
                           ((aDroid(psDroid).action = DACTION_MOVETOREARM) Or _
                            (aDroid(psDroid).action = DACTION_WAITFORREARM)) And _
                           (aDroid(psDroid).sMove.Status <> MOVEINACTIVE) Then
                        ' catch vtols that were attacking another target which was destroyed
                        ' get them to attack the new target rather than returning to rearm
                        bAttack = True
                    ElseIf Not vtolEmpty(psDroid) Then
                        bAttack = True
                    End If
                Else
                    bAttack = True
                End If

                'if not currently attacking or target has changed
                ' (&& and ||: isAttacking, whose CHECK_DROID may log, only when attacking)
                c = bAttack
                If c Then
                    c = Not DROID_isAttacking(psDroid)
                    If Not c Then c = psFireTarget <> aDroid(psDroid).psActionTarget(0)
                End If
                If c Then
                    'get the droid to attack
                    actionDroid_obj psDroid, DACTION_ATTACK, psFireTarget
                End If
            ElseIf DROID_isVtol(psDroid) And _
                   (aDroid(psDroid).action <> DACTION_NONE) And _
                   (aDroid(psDroid).action <> DACTION_FIRESUPPORT) Then
                moveToRearm psDroid
            ElseIf (aDroid(psDroid).action <> DACTION_FIRESUPPORT) And _
                   (aDroid(psDroid).action <> DACTION_FIRESUPPORT_RETREAT) Then
                actionDroid_obj psDroid, DACTION_FIRESUPPORT, aDroid(psDroid).order.psObj
            End If
        End If
    Case DORDER_RECYCLE
        ' don't bother with formations for this order
        If aDroid(psDroid).sMove.psFormation <> 0 Then
            formationLeave aDroid(psDroid).sMove.psFormation, psDroid
            aDroid(psDroid).sMove.psFormation = 0
        End If
        If aDroid(psDroid).order.psObj = 0 Then
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            actionDroid psDroid, DACTION_NONE
        ElseIf actionReachedBuildPos(psDroid, aStruct(aDroid(psDroid).order.psObj).pos.x, aStruct(aDroid(psDroid).order.psObj).pos.y, aStruct(aDroid(psDroid).order.psObj).rot.direction, structureStatsBase(aStruct(aDroid(psDroid).order.psObj).pStructureType)) Then
            recycleDroid psDroid
        ElseIf aDroid(psDroid).action = DACTION_NONE Then
            objPos = SIMPLE_OBJECT_pos(aDroid(psDroid).order.psObj)
            actionDroid_loc psDroid, DACTION_MOVE, objPos.x, objPos.y
        End If
    Case DORDER_GUARD
        If orderDroidList(psDroid) Then
            ' started a queued order - quit
            GoTo switchEnd
        ElseIf (aDroid(psDroid).action = DACTION_NONE) Or _
               (aDroid(psDroid).action = DACTION_MOVE) Or _
               (aDroid(psDroid).action = DACTION_MOVEFIRE) Then
            ' not doing anything, make sure the droid is close enough
            ' to the thing it is defending
            If guardsCommander(psDroid) And Not (aDroid(psDroid).droidType = DROID_REPAIR Or aDroid(psDroid).droidType = DROID_CYBORG_REPAIR) Then
                ' guarding a commander, allow more space
                orderCheckGuardPosition psDroid, DEFEND_CMD_BASEDIST
            Else
                orderCheckGuardPosition psDroid, DEFEND_BASEDIST
            End If
        ElseIf aDroid(psDroid).droidType = DROID_REPAIR Or aDroid(psDroid).droidType = DROID_CYBORG_REPAIR Then
            ' repairing something, make sure the droid doesn't go too far
            orderCheckGuardPosition psDroid, REPAIR_MAXDIST
        ElseIf aDroid(psDroid).droidType = DROID_CONSTRUCT Or aDroid(psDroid).droidType = DROID_CYBORG_CONSTRUCT Then
            ' repairing something, make sure the droid doesn't go too far
            orderCheckGuardPosition psDroid, CONSTRUCT_MAXDIST
        ElseIf DROID_isTransporter(psDroid) Then

        Else
            'let vtols return to rearm
            If Not DROID_isVtolRearming(psDroid) Then
                ' attacking something, make sure the droid doesn't go too far
                If guardsCommander(psDroid) Then
                    ' guarding a commander, allow more space
                    orderCheckGuardPosition psDroid, DEFEND_CMD_MAXDIST
                Else
                    orderCheckGuardPosition psDroid, DEFEND_MAXDIST
                End If
            End If
        End If

        ' get combat units in a command group to attack the commanders target
        If hasCommander(psDroid) And (aDroid(psDroid).numWeaps > 0) Then
            commander = aGroup(aDroid(psDroid).psGroup).psCommander
            c = False
            If aDroid(commander).action = DACTION_ATTACK And _
               aDroid(commander).psActionTarget(0) <> 0 Then
                c = SIMPLE_OBJECT_died(aDroid(commander).psActionTarget(0)) = 0
            End If
            If c Then
                psObj = aDroid(commander).psActionTarget(0)
                If aDroid(psDroid).action = DACTION_ATTACK Or _
                   aDroid(psDroid).action = DACTION_MOVETOATTACK Then
                    If aDroid(psDroid).psActionTarget(0) <> psObj Then
                        actionDroid_obj psDroid, DACTION_ATTACK, psObj
                    End If
                ElseIf aDroid(psDroid).action <> DACTION_MOVE Then
                    actionDroid_obj psDroid, DACTION_ATTACK, psObj
                End If
            End If

            ' make sure units in a command group are actually guarding the commander
            psObj = orderStateObj(psDroid, DORDER_GUARD) ' find out who is being guarded by the droid
            If psObj = 0 _
               Or psObj <> aGroup(aDroid(psDroid).psGroup).psCommander Then
                orderDroidObj psDroid, DORDER_GUARD, aGroup(aDroid(psDroid).psGroup).psCommander, ModeImmediate
            End If
        End If

        tryDoRepairlikeAction psDroid
    Case Else
        ASSERT_FAILURE "false", "src/order.cpp:1238", "orderUpdateDroid", "orderUpdateUnit: unknown order"
    End Select
switchEnd:

    ' catch any vtol that is rearming but has finished his order
    ' (&&: isVtolRearming, whose CHECK_DROID may log, only for a droid without an order)
    If aDroid(psDroid).order.typ = DORDER_NONE Then
        If DROID_isVtolRearming(psDroid) Then
            c = aDroid(psDroid).psActionTarget(0) = 0
            If Not c Then c = SIMPLE_OBJECT_died(aDroid(psDroid).psActionTarget(0)) = 0
            If c Then
                aDroid(psDroid).order = DroidOrder_obj(DORDER_REARM, aDroid(psDroid).psActionTarget(0))
            End If
        End If
    End If

    If aDroid(psDroid).selected <> 0 Then
        ' Tell us what the droid is doing.
        DROIDDOING = va("%.12s,id(%d) order(%d):%s action(%d):%s secondary:%x move:%s", droidGetName(psDroid), aDroid(psDroid).id, _
                        aDroid(psDroid).order.typ, getDroidOrderName(aDroid(psDroid).order.typ), aDroid(psDroid).action, getDroidActionName(aDroid(psDroid).action), aDroid(psDroid).secondaryOrder, _
                        moveDescription(aDroid(psDroid).sMove.Status))
    End If

    orderUpdateDroid = True
End Function

' (VB) orderUpdateDroid, DORDER_GUARD: psDroid->order.psObj != nullptr && psDroid->order.psObj->type == OBJ_DROID &&
' ((DROID *)psDroid->order.psObj)->droidType == DROID_COMMAND
Private Function guardsCommander(ByVal psDroid As Long) As Boolean
    If aDroid(psDroid).order.psObj <> 0 Then
        If SIMPLE_OBJECT_type(aDroid(psDroid).order.psObj) = OBJ_DROID Then
            guardsCommander = aDroid(aDroid(psDroid).order.psObj).droidType = DROID_COMMAND
        End If
    End If
End Function

' This function sends all members of the psGroup the order psData using orderDroidBase().
' If the order data is to recover an artifact, the order is only given to the closest droid to the artifact.
Private Sub orderCmdGroupBase(ByVal psGroup As Long, psData As DROID_ORDER_DATA)
    Dim psChosen As Long, mindist As Long, it As Long, psCurr As Long, currdist As Long
    Dim isAttackOrder As Boolean, observeOrder As DROID_ORDER_DATA, i As Long

    If Not (psGroup <> 0) Then ASSERT_FAILURE "psGroup != nullptr", "src/order.cpp:1265", "orderCmdGroupBase", "Invalid unit group": Exit Sub

    syncDebug "orderCmdGroupBase", "Commander group order"

    If psData.typ = DORDER_RECOVER Then
        ' picking up an artifact - only need to send one unit
        psChosen = 0
        mindist = SDWORD_MAX
        it = aGroup(psGroup).psList.head
        Do While it <> 0
            psCurr = lnObj(it)
            If aDroid(psCurr).order.typ = DORDER_RTR Or aDroid(psCurr).order.typ = DORDER_RTB Or aDroid(psCurr).order.typ = DORDER_RTR_SPECIFIED Then
                ' don't touch units returning for repairs
                GoTo continueRecover
            End If
            currdist = objPosDiffSq(aDroid(psCurr).pos, SIMPLE_OBJECT_pos(psData.psObj))
            If currdist < mindist Then
                psChosen = psCurr
                mindist = currdist
            End If
            syncDebug "orderCmdGroupBase", "command %d,%d", aDroid(psCurr).id, currdist
continueRecover:
            it = lnNext(it)
        Loop
        If psChosen <> 0 Then
            orderDroidBase psChosen, psData
        End If
    Else
        isAttackOrder = psData.typ = DORDER_ATTACKTARGET Or psData.typ = DORDER_ATTACK
        it = aGroup(psGroup).psList.head
        Do While it <> 0
            psCurr = lnObj(it)
            syncDebug "orderCmdGroupBase", "command %d", aDroid(psCurr).id
            If Not orderState(psCurr, DORDER_RTR) Then  ' if you change this, youll need to change sendcmdgroup()
                If Not isAttackOrder Then
                    orderDroidBase psCurr, psData
                    GoTo continueOrder
                End If
                If aDroid(psCurr).droidType = DROID_SENSOR And psData.psObj <> 0 Then
                    ' sensors must observe, not attack
                    observeOrder = DroidOrder_obj(DORDER_OBSERVE, psData.psObj)
                    orderDroidBase psCurr, observeOrder
                Else
                    ' for non-sensors, check that the designated target is actually valid
                    ' there is no point in ordering AA gun to attack ground units
                    For i = 0 To MAX_WEAPONS - 1

                        If validTarget(psCurr, psData.psObj, i) Then
                            orderDroidBase psCurr, psData
                            Exit For
                        End If
                    Next
                End If
            End If
continueOrder:
            it = lnNext(it)
        Loop
    End If
End Sub

' This function chooses the sound to play after the object is assigned to fire support a specific unit. Uses audio_QueueTrackMinDelay() to play the sound.
' @todo this function is about playing audio. I'm not sure this should be in here.
Private Sub orderPlayFireSupportAudio(ByVal psObj As Long)
    Dim iAudioID As Long, psDroid As Long, psStruct As Long

    iAudioID = NO_SOUND

    If Not (psObj <> 0) Then ASSERT_FAILURE "psObj != nullptr", "src/order.cpp:1346", "orderPlayFireSupportAudio", "Invalid pointer": Exit Sub
    ' play appropriate speech
    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        psDroid = psObj
        If aDroid(psDroid).droidType = DROID_COMMAND Then
            iAudioID = ID_SOUND_ASSIGNED_TO_COMMANDER
        ElseIf aDroid(psDroid).droidType = DROID_SENSOR Then
            iAudioID = ID_SOUND_ASSIGNED_TO_SENSOR
        End If

    Case OBJ_STRUCTURE
        psStruct = psObj
        'check for non-CB first
        If structStandardSensor(psStruct) Or structVTOLSensor(psStruct) Then
            iAudioID = ID_SOUND_ASSIGNED_TO_SENSOR
        ElseIf structCBSensor(psStruct) Or structVTOLCBSensor(psStruct) Then
            iAudioID = ID_SOUND_ASSIGNED_TO_COUNTER_RADAR
        End If
    Case Else
    End Select

    If iAudioID <> NO_SOUND Then
        audio_QueueTrackMinDelay iAudioID, AUDIO_DELAY_FIRESUPPORT
    End If
End Sub

' This function actually tells the droid to perform the psOrder.
' This function is called everytime to send a direct order to a droid.
Public Sub orderDroidBase(ByVal psDroid As Long, psOrder As DROID_ORDER_DATA)
    Dim iFactoryDistSq As Long, psFactory As Long, psPropStats As Long, rPos As Vector3i
    Dim oldState As Long, it As Long, psStruct As Long, pos As Vector2i, iDX As Long, iDY As Long, startPos As Vector2i
    Dim rtrData As RtrBestResult, embarkee As Long, iStructDistSq As Long, psStats As Long, c As Boolean
    Dim objPos As Vector3i, statsName As String

    psPropStats = DROID_getPropulsionStats(psDroid)
    rPos = Vector3i_(psOrder.pos.x, psOrder.pos.y, 0)
    syncDebugDroid_ "orderDroidBase", psDroid, Asc("-")
    syncDebug "orderDroidBase", "%d ordered %s", aDroid(psDroid).id, getDroidOrderName(psOrder.typ)

    objTrace aDroid(psDroid).id, "orderDroidBase", "base set order to %s (was %s)", getDroidOrderName(psOrder.typ), getDroidOrderName(aDroid(psDroid).order.typ)

    If psOrder.typ <> DORDER_TRANSPORTIN _
       And psOrder.psObj = 0 _
       And (validOrderForLoc(psOrder.typ) Or psOrder.typ = DORDER_BUILD) Then
        If Not fpathCheck(aDroid(psDroid).pos, rPos, asPropulsionStats(psPropStats).propulsionType) Then
            If Not isHumanPlayer(aDroid(psDroid).player) Then
                debug_ 1406, LOG_SCRIPT, "orderDroidBase", "Invalid order %s given to player %d's %s for position (%d, %d) - ignoring", _
                       getDroidOrderName(psOrder.typ), aDroid(psDroid).player, droidGetName(psDroid), psOrder.pos.x, psOrder.pos.y
            End If
            objTrace aDroid(psDroid).id, "orderDroidBase", "Invalid order %s for position (%d, %d) - ignoring", getDroidOrderName(psOrder.typ), psOrder.pos.x, psOrder.pos.y
            syncDebugDroid_ "orderDroidBase", psDroid, Asc("?")
            Exit Sub
        End If
    End If

    ' deal with a droid receiving a primary order
    If Not DROID_isTransporter(psDroid) And psOrder.typ <> DORDER_NONE And psOrder.typ <> DORDER_STOP And psOrder.typ <> DORDER_GUARD Then
        ' reset secondary order
        oldState = aDroid(psDroid).secondaryOrder
        aDroid(psDroid).secondaryOrder = aDroid(psDroid).secondaryOrder And Not (DSS_RTL_MASK Or DSS_RECYCLE_MASK Or DSS_PATROL_MASK)
        aDroid(psDroid).secondaryOrderPending = aDroid(psDroid).secondaryOrderPending And Not (DSS_RTL_MASK Or DSS_RECYCLE_MASK Or DSS_PATROL_MASK)
        objTrace aDroid(psDroid).id, "orderDroidBase", "secondary order reset due to primary order set"
        If oldState <> aDroid(psDroid).secondaryOrder And aDroid(psDroid).player = selectedPlayer Then
            intRefreshScreen
        End If
    End If

    ' if this is a command droid - all its units do the same thing
    c = False
    If (aDroid(psDroid).droidType = DROID_COMMAND) And _
       (aDroid(psDroid).psGroup <> 0) Then
        c = (aGroup(aDroid(psDroid).psGroup).typ = GT_COMMAND) And _
            (psOrder.typ <> DORDER_GUARD) And _
            (psOrder.typ <> DORDER_RTR) And _
            (psOrder.typ <> DORDER_RECYCLE)  '(psOrder->psObj == NULL)) &&
    End If
    If c Then
        If psOrder.typ = DORDER_ATTACK Then
            ' change to attacktarget so that the group members
            ' guard order does not get canceled
            psOrder.typ = DORDER_ATTACKTARGET
            orderCmdGroupBase aDroid(psDroid).psGroup, psOrder
            psOrder.typ = DORDER_ATTACK
        Else
            orderCmdGroupBase aDroid(psDroid).psGroup, psOrder
        End If

        ' the commander doesn't have to pick up artifacts, one
        ' of his units will do it for him (if there are any in his group).
        If (psOrder.typ = DORDER_RECOVER) And _
           (aGroup(aDroid(psDroid).psGroup).psList.size <> 0) Then
            psOrder.typ = DORDER_NONE
        End If
    End If

    ' A selected campaign transporter shouldn't be given orders by the player.
    ' Campaign transporter selection is required for it to be tracked by the camera, and
    ' should be the only case when it does get selected.
    If DROID_isTransporter(psDroid) And _
       Not bMultiPlayer And _
       aDroid(psDroid).selected <> 0 And _
       (psOrder.typ <> DORDER_TRANSPORTOUT And _
        psOrder.typ <> DORDER_TRANSPORTIN And _
        psOrder.typ <> DORDER_TRANSPORTRETURN) Then
        Exit Sub
    End If

    Select Case psOrder.typ
    Case DORDER_NONE
        ' used when choose order cannot assign an order
    Case DORDER_STOP
        ' get the droid to stop doing whatever it is doing
        actionDroid psDroid, DACTION_NONE
        aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
    Case DORDER_HOLD
        ' get the droid to stop doing whatever it is doing and temp hold
        actionDroid psDroid, DACTION_NONE
        aDroid(psDroid).order = psOrder
    Case DORDER_MOVE, DORDER_SCOUT
        ' can't move vtols to blocking tiles
        If DROID_isVtol(psDroid) Then
            If fpathBlockingTile(map_coord(psOrder.pos.x), map_coord(psOrder.pos.y), asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType) Then
                GoTo switchEnd
            End If
        End If
        'in multiPlayer, cannot move Transporter to blocking tile either
        If game.typ = LEVEL_TYPE_SKIRMISH _
           And DROID_isTransporter(psDroid) Then
            If fpathBlockingTile(map_coord(psOrder.pos.x), map_coord(psOrder.pos.y), asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType) Then
                GoTo switchEnd
            End If
        End If
        ' move a droid to a location
        aDroid(psDroid).order = psOrder
        actionDroid_loc psDroid, DACTION_MOVE, psOrder.pos.x, psOrder.pos.y
    Case DORDER_PATROL
        aDroid(psDroid).order = psOrder
        aDroid(psDroid).order.pos2.x = aDroid(psDroid).pos.x
        aDroid(psDroid).order.pos2.y = aDroid(psDroid).pos.y
        actionDroid_loc psDroid, DACTION_MOVE, psOrder.pos.x, psOrder.pos.y
    Case DORDER_RECOVER
        aDroid(psDroid).order = psOrder
        objPos = SIMPLE_OBJECT_pos(psOrder.psObj)
        actionDroid_loc psDroid, DACTION_MOVE, objPos.x, objPos.y
    Case DORDER_TRANSPORTOUT
        ' tell a (transporter) droid to leave home base for the offworld mission
        aDroid(psDroid).order = psOrder
        actionDroid_loc psDroid, DACTION_TRANSPORTOUT, psOrder.pos.x, psOrder.pos.y
    Case DORDER_TRANSPORTRETURN
        ' tell a (transporter) droid to return after unloading
        aDroid(psDroid).order = psOrder
        actionDroid_loc psDroid, DACTION_TRANSPORTOUT, psOrder.pos.x, psOrder.pos.y
    Case DORDER_TRANSPORTIN
        ' tell a (transporter) droid to fly onworld
        aDroid(psDroid).order = psOrder
        actionDroid_loc psDroid, DACTION_TRANSPORTIN, psOrder.pos.x, psOrder.pos.y
    Case DORDER_ATTACK, DORDER_ATTACKTARGET
        If aDroid(psDroid).numWeaps = 0 _
           Or aDroid(psDroid).asWeaps(0).nStat = 0 _
           Or DROID_isTransporter(psDroid) Then
            GoTo switchEnd
        ElseIf aDroid(psDroid).order.typ = DORDER_GUARD And psOrder.typ = DORDER_ATTACKTARGET Then
            ' attacking something while guarding, don't change the order
            actionDroid_obj psDroid, DACTION_ATTACK, psOrder.psObj
        ElseIf psOrder.psObj <> 0 Then
            If SIMPLE_OBJECT_died(psOrder.psObj) = 0 Then
                'cannot attack a Transporter with EW in multiPlayer
                ' FIXME: Why not ?
                c = False
                If game.typ = LEVEL_TYPE_SKIRMISH Then
                    If electronicDroid(psDroid) Then
                        If SIMPLE_OBJECT_type(psOrder.psObj) = OBJ_DROID Then c = DROID_isTransporter(psOrder.psObj)
                    End If
                End If
                If c Then
                    GoTo switchEnd
                End If
                aDroid(psDroid).order = psOrder

                c = DROID_isVtol(psDroid)
                If Not c Then c = actionInRange(psDroid, psOrder.psObj, 0)
                If Not c Then c = ((psOrder.typ = DORDER_ATTACKTARGET Or psOrder.typ = DORDER_ATTACK) And secondaryGetState(psDroid, DSO_HALTTYPE) = DSS_HALT_HOLD)
                If c Then
                    ' when DSS_HALT_HOLD, don't move to attack
                    actionDroid_obj psDroid, DACTION_ATTACK, psOrder.psObj
                Else
                    objPos = SIMPLE_OBJECT_pos(psOrder.psObj)
                    actionDroid_loc psDroid, DACTION_MOVE, objPos.x, objPos.y
                End If
            End If
        End If
    Case DORDER_BUILD, DORDER_LINEBUILD
        ' build a new structure or line of structures
        If Not (DROID_isConstructionDroid(psDroid)) Then ASSERT_FAILURE "psDroid->isConstructionDroid()", "src/order.cpp:1569", "orderDroidBase", "%s cannot construct things!", objInfo(psDroid): Exit Sub
        If Not (psOrder.psStats <> -1) Then ASSERT_FAILURE "psOrder->psStats != nullptr", "src/order.cpp:1570", "orderDroidBase", "invalid structure stats pointer": Exit Sub
        aDroid(psDroid).order = psOrder
        c = aDroid(psDroid).order.psStats = -1
        If Not c Then c = asStructureStats(aDroid(psDroid).order.psStats).typ <> REF_DEMOLISH
        If Not c Then ASSERT_FAILURE "!psDroid->order.psStats || psDroid->order.psStats->type != REF_DEMOLISH", "src/order.cpp:1572", "orderDroidBase", "Cannot build demolition": Exit Sub
        actionDroid_loc psDroid, DACTION_BUILD, psOrder.pos.x, psOrder.pos.y
        If psOrder.psStats <> -1 Then statsName = getStatsName_(psOrder.psStats) Else statsName = "NULL"
        objTrace aDroid(psDroid).id, "orderDroidBase", "Starting new construction effort of %s", statsName
    Case DORDER_BUILDMODULE
        'build a module onto the structure
        c = Not DROID_isConstructionDroid(psDroid)
        If Not c Then c = ULT(psOrder.index, nextModuleToBuild(psOrder.psObj, -1))
        If c Then
            GoTo switchEnd
        End If
        psStats = getModuleStat(psOrder.psObj)
        If Not (psStats <> -1) Then ASSERT_FAILURE "psStats != nullptr", "src/order.cpp:1584", "orderDroidBase", "should have found a module stats": Exit Sub
        objPos = SIMPLE_OBJECT_pos(psOrder.psObj)
        aDroid(psDroid).order = DroidOrder_stats(DORDER_BUILD, psStats, Vector2i_(objPos.x, objPos.y), 0)
        c = aDroid(psDroid).order.psStats = -1
        If Not c Then c = asStructureStats(aDroid(psDroid).order.psStats).typ <> REF_DEMOLISH
        If Not c Then ASSERT_FAILURE "!psDroid->order.psStats || psDroid->order.psStats->type != REF_DEMOLISH", "src/order.cpp:1586", "orderDroidBase", "Cannot build demolition": Exit Sub
        actionDroid_loc psDroid, DACTION_BUILD, objPos.x, objPos.y
        If psOrder.psStats <> -1 Then statsName = getStatsName_(psOrder.psStats) Else statsName = "NULL"
        objTrace aDroid(psDroid).id, "orderDroidBase", "Starting new upgrade of %s", statsName
    Case DORDER_HELPBUILD
        ' help to build a structure that is starting to be built
        If Not (DROID_isConstructionDroid(psDroid)) Then ASSERT_FAILURE "psDroid->isConstructionDroid()", "src/order.cpp:1593", "orderDroidBase", "Not a constructor droid": Exit Sub
        If Not (psOrder.psObj <> 0) Then ASSERT_FAILURE "psOrder->psObj != nullptr", "src/order.cpp:1594", "orderDroidBase", "Help to build a NULL pointer?": Exit Sub
        If aDroid(psDroid).action = DACTION_BUILD And psOrder.psObj = aDroid(psDroid).psActionTarget(0) _
           And (aDroid(psDroid).order.typ <> DORDER_LINEBUILD) Then
            ' skip DORDER_LINEBUILD -> we still want to drop pending structure blueprints
            ' this isn't a perfect solution, because ordering a LINEBUILD with negative energy, and then clicking
            ' on first structure being built, will remove it, as we change order from DORDR_LINEBUILD to DORDER_BUILD
            ' we are already building it, nothing to do
            objTrace aDroid(psDroid).id, "orderDroidBase", "Ignoring DORDER_HELPBUILD because already buildig object %i", SIMPLE_OBJECT_id(psOrder.psObj)
            GoTo switchEnd
        End If
        aDroid(psDroid).order = psOrder
        objPos = SIMPLE_OBJECT_pos(psOrder.psObj)
        aDroid(psDroid).order.pos = Vector2i_(objPos.x, objPos.y)
        aDroid(psDroid).order.psStats = aStruct(psOrder.psObj).pStructureType
        c = aDroid(psDroid).order.psStats = -1
        If Not c Then c = asStructureStats(aDroid(psDroid).order.psStats).typ <> REF_DEMOLISH
        If Not c Then ASSERT_FAILURE "!psDroid->order.psStats || psDroid->order.psStats->type != REF_DEMOLISH", "src/order.cpp:1608", "orderDroidBase", "Cannot build demolition": Exit Sub
        actionDroid_loc psDroid, DACTION_BUILD, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
        If psOrder.psStats <> -1 Then statsName = getStatsName_(aDroid(psDroid).order.psStats) Else statsName = "NULL"
        objTrace aDroid(psDroid).id, "orderDroidBase", "Helping construction of %s", statsName
    Case DORDER_DEMOLISH
        If Not (aDroid(psDroid).droidType = DROID_CONSTRUCT Or aDroid(psDroid).droidType = DROID_CYBORG_CONSTRUCT) Then
            GoTo switchEnd
        End If
        aDroid(psDroid).order = psOrder
        objPos = SIMPLE_OBJECT_pos(psOrder.psObj)
        aDroid(psDroid).order.pos = Vector2i_(objPos.x, objPos.y)
        actionDroid_obj psDroid, DACTION_DEMOLISH, psOrder.psObj
    Case DORDER_REPAIR
        If Not (aDroid(psDroid).droidType = DROID_CONSTRUCT Or aDroid(psDroid).droidType = DROID_CYBORG_CONSTRUCT) Then
            GoTo switchEnd
        End If
        aDroid(psDroid).order = psOrder
        objPos = SIMPLE_OBJECT_pos(psOrder.psObj)
        aDroid(psDroid).order.pos = Vector2i_(objPos.x, objPos.y)
        actionDroid_obj psDroid, DACTION_REPAIR, psOrder.psObj
    Case DORDER_DROIDREPAIR
        If Not (aDroid(psDroid).droidType = DROID_REPAIR Or aDroid(psDroid).droidType = DROID_CYBORG_REPAIR) Then
            GoTo switchEnd
        End If
        aDroid(psDroid).order = psOrder
        actionDroid_obj psDroid, DACTION_DROIDREPAIR, psOrder.psObj
    Case DORDER_OBSERVE
        ' keep an object within sensor view
        aDroid(psDroid).order = psOrder
        actionDroid_obj psDroid, DACTION_OBSERVE, psOrder.psObj
    Case DORDER_FIRESUPPORT
        If DROID_isTransporter(psDroid) Then
            debug_ 1646, LOG_ERROR, "orderDroidBase", "Sorry, transports cannot be assigned to commanders."
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            GoTo switchEnd
        End If
        If aDroid(psDroid).asWeaps(0).nStat = 0 Then
            GoTo switchEnd
        End If
        aDroid(psDroid).order = psOrder
        ' let the order update deal with vtol droids
        If Not DROID_isVtol(psDroid) Then
            actionDroid_obj psDroid, DACTION_FIRESUPPORT, psOrder.psObj
        End If

        If aDroid(psDroid).player = selectedPlayer Then
            orderPlayFireSupportAudio psOrder.psObj
        End If
    Case DORDER_COMMANDERSUPPORT
        If DROID_isTransporter(psDroid) Then
            debug_ 1669, LOG_ERROR, "orderDroidBase", "Sorry, transports cannot be assigned to commanders."
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            GoTo switchEnd
        End If
        If Not (psOrder.psObj <> 0) Then ASSERT_FAILURE "psOrder->psObj != nullptr", "src/order.cpp:1673", "orderDroidBase", "Can't command a NULL": Exit Sub
        If cmdDroidAddDroid(psOrder.psObj, psDroid) And aDroid(psDroid).player = selectedPlayer Then
            orderPlayFireSupportAudio psOrder.psObj
        ElseIf aDroid(psDroid).player = selectedPlayer Then
            audio_PlayBuildFailedOnce
        End If
    Case DORDER_RTB
        it = apsStructLists(aDroid(psDroid).player).head
        Do While it <> 0
            psStruct = lnObj(it)
            If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_HQ Then
                pos.x = aStruct(psStruct).pos.x
                pos.y = aStruct(psStruct).pos.y
                If Not CheckInScrollLimits(pos.x, pos.y) Then
                    GoTo continueRtb
                End If

                aDroid(psDroid).order = psOrder
                ' Find a place to land for vtols. And Transporters in a multiPlay game.
                If DROID_isVtol(psDroid) Or (game.typ = LEVEL_TYPE_SKIRMISH And DROID_isTransporter(psDroid)) Then
                    actionVTOLLandingPos psDroid, pos
                End If
                actionDroid_loc psDroid, DACTION_MOVE, pos.x, pos.y
                If aDroid(psDroid).sMove.psFormation <> 0 Then
                    formationLeave aDroid(psDroid).sMove.psFormation, psDroid
                    aDroid(psDroid).sMove.psFormation = 0
                End If
                Exit Do
            End If
continueRtb:
            it = lnNext(it)
        Loop
        ' no HQ go to the landing zone
        If aDroid(psDroid).order.typ <> DORDER_RTB Then
            ' see if the LZ has been set up
            iDX = getLandingX(aDroid(psDroid).player)
            iDY = getLandingY(aDroid(psDroid).player)
            startPos = getPlayerStartPosition(aDroid(psDroid).player)

            c = False
            If iDX <> 0 And iDY <> 0 Then c = CheckInScrollLimits(iDX, iDY)
            If c Then
                aDroid(psDroid).order = psOrder
                actionDroid_loc psDroid, DACTION_MOVE, iDX, iDY
            ElseIf bMultiPlayer And (startPos.x <> 0 And startPos.y <> 0) Then
                If CheckInScrollLimits(startPos.x, startPos.y) Then
                    aDroid(psDroid).order = psOrder
                    actionDroid_loc psDroid, DACTION_MOVE, startPos.x, startPos.y
                End If
            End If
        End If
    Case DORDER_RTR_SPECIFIED
        ' DORDER_RTR_SPECIFIED requires special handling, different from RTR
        ' because we don't want the damaged droid to randomly chose a repair station
        ' when User already specified one explicitely
        If DROID_isVtol(psDroid) Then
            moveToRearm psDroid
            GoTo switchEnd
        End If
        ' if already has a target repair, don't override it: it might be different
        ' and we don't want come back and forth between 2 repair points
        If aDroid(psDroid).order.typ = DORDER_RTR_SPECIFIED And psOrder.psObj <> 0 Then
            If SIMPLE_OBJECT_died(psOrder.psObj) = 0 Then
                objTrace aDroid(psDroid).id, "orderDroidBase", "DONE FOR NOW"
                GoTo switchEnd
            End If
        End If

        ' RTR_SPECIFIED can only ever be given to a repair facility, never a mobile repair turret
        ' move to front of structure
        aDroid(psDroid).order = DroidOrder_objRtr(psOrder.typ, psOrder.psObj, RTR_TYPE_REPAIR_FACILITY)
        objPos = SIMPLE_OBJECT_pos(psOrder.psObj)
        aDroid(psDroid).order.pos = Vector2i_(objPos.x, objPos.y)
        ' If in multiPlayer, and the Transporter has been sent to be
        ' repaired, need to find a suitable location to drop down.
        If game.typ = LEVEL_TYPE_SKIRMISH And DROID_isTransporter(psDroid) Then
            pos = aDroid(psDroid).order.pos
            objTrace aDroid(psDroid).id, "orderDroidBase", "Repair transport"
            actionVTOLLandingPos psDroid, pos
            actionDroid_loc psDroid, DACTION_MOVE, pos.x, pos.y
        Else
            objTrace aDroid(psDroid).id, "orderDroidBase", "RTR_SPECIFIED: Go to repair facility at (%d, %d)!", aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
            actionDroid_objLoc psDroid, DACTION_MOVE, psOrder.psObj, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
        End If
    Case DORDER_RTR
        If DROID_isVtol(psDroid) Then
            moveToRearm psDroid
            GoTo switchEnd
        End If
        ' if already has a target repair, don't override it: it might be different
        ' and we don't want come back and forth between 2 repair points
        If aDroid(psDroid).order.typ = DORDER_RTR And psOrder.psObj <> 0 Then
            If SIMPLE_OBJECT_died(psOrder.psObj) = 0 Then
                objTrace aDroid(psDroid).id, "orderDroidBase", "DONE FOR NOW"
                GoTo switchEnd
            End If
        End If
        rtrData = RtrBestResult_(RTR_TYPE_NO_RESULT, 0)
        If psOrder.typ = DORDER_RTR And (psOrder.rtrType = RTR_TYPE_NO_RESULT Or psOrder.psObj = 0) Then
            rtrData = decideWhereToRepairAndBalance(psDroid)
        Else
            rtrData = RtrBestResult_order(psOrder)
        End If

        ' give repair order if repair facility found
        If rtrData.typ = RTR_TYPE_REPAIR_FACILITY Then
            ' move to front of structure
            aDroid(psDroid).order = DroidOrder_objRtr(psOrder.typ, rtrData.psObj, RTR_TYPE_REPAIR_FACILITY)
            objPos = SIMPLE_OBJECT_pos(rtrData.psObj)
            aDroid(psDroid).order.pos = Vector2i_(objPos.x, objPos.y)
            ' If in multiPlayer, and the Transporter has been sent to be
            ' repaired, need to find a suitable location to drop down.
            If game.typ = LEVEL_TYPE_SKIRMISH And DROID_isTransporter(psDroid) Then
                pos = aDroid(psDroid).order.pos

                objTrace aDroid(psDroid).id, "orderDroidBase", "Repair transport"
                actionVTOLLandingPos psDroid, pos
                actionDroid_loc psDroid, DACTION_MOVE, pos.x, pos.y
            Else
                objTrace aDroid(psDroid).id, "orderDroidBase", "Go to repair facility at (%d, %d)!", objPos.x, objPos.y
                actionDroid_objLoc psDroid, DACTION_MOVE, rtrData.psObj, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
            End If
        ' give repair order if repair droid found
        ElseIf rtrData.typ = RTR_TYPE_DROID And Not DROID_isTransporter(psDroid) Then
            objPos = SIMPLE_OBJECT_pos(rtrData.psObj)
            aDroid(psDroid).order = DroidOrder_posRtr(psOrder.typ, Vector2i_(objPos.x, objPos.y), RTR_TYPE_DROID)
            aDroid(psDroid).order.pos = Vector2i_(objPos.x, objPos.y)
            aDroid(psDroid).order.psObj = rtrData.psObj
            objTrace aDroid(psDroid).id, "orderDroidBase", "Go to repair at (%d, %d) using (%d, %d), time %i!", objPos.x, objPos.y, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y, gameTime
            actionDroid_loc psDroid, DACTION_MOVE, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
        Else
            ' no repair facility or HQ go to the landing zone
            If Not bMultiPlayer And selectedPlayer = 0 Then
                objTrace aDroid(psDroid).id, "orderDroidBase", "could not RTR, doing RTL instead"
                orderDroid psDroid, DORDER_RTB, ModeImmediate
            End If
        End If
    Case DORDER_EMBARK
        embarkee = castDroid(psOrder.psObj)
        c = DROID_isTransporter(psDroid)  ' require a transporter for embarking.
        If Not c Then c = embarkee = 0
        If Not c Then c = Not DROID_isTransporter(embarkee)  ' nor can a transporter load another transporter
        If c Then
            debug_ 1839, LOG_ERROR, "orderDroidBase", "Sorry, can only load things that aren't transporters into things that are."
            aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
            GoTo switchEnd
        End If
        ' move the droid to the transporter location
        aDroid(psDroid).order = psOrder
        objPos = SIMPLE_OBJECT_pos(psOrder.psObj)
        aDroid(psDroid).order.pos = Vector2i_(objPos.x, objPos.y)
        actionDroid_loc psDroid, DACTION_MOVE, objPos.x, objPos.y
    Case DORDER_DISEMBARK
        'only valid in multiPlayer mode
        If bMultiPlayer Then
            'this order can only be given to Transporter droids
            If DROID_isTransporter(psDroid) Then
                aDroid(psDroid).order = psOrder
                'move the Transporter to the requested location
                actionDroid_loc psDroid, DACTION_MOVE, psOrder.pos.x, psOrder.pos.y
                'close the Transporter interface - if up
                If Not widgGetFromID(psWScreen, IDTRANS_FORM) Is Nothing Then
                    intRemoveTrans
                End If
            End If
        End If
    Case DORDER_RECYCLE
        psFactory = 0
        iFactoryDistSq = 0
        it = apsStructLists(aDroid(psDroid).player).head
        Do While it <> 0
            psStruct = lnObj(it)
            ' Look for nearest factory or repair facility
            Select Case asStructureStats(aStruct(psStruct).pStructureType).typ
            Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY, REF_REPAIR_FACILITY
                ' get droid->facility distance squared
                iStructDistSq = droidSqDist(psDroid, psStruct)

                ' Choose current structure if first facility found or nearer than previously chosen facility
                If aStruct(psStruct).status = SS_BUILT And iStructDistSq > 0 Then
                    If psFactory = 0 Then
                        psFactory = psStruct
                        iFactoryDistSq = iStructDistSq
                    ElseIf ULT(iStructDistSq, iFactoryDistSq) Then
                        psFactory = psStruct
                        iFactoryDistSq = iStructDistSq
                    End If
                End If
            End Select
            it = lnNext(it)
        Loop

        ' give recycle order if facility found
        If psFactory <> 0 Then
            ' move to front of structure
            aDroid(psDroid).order = DroidOrder_obj(psOrder.typ, psFactory)
            aDroid(psDroid).order.pos = Vector2i_(aStruct(psFactory).pos.x, aStruct(psFactory).pos.y)
            setDroidTarget psDroid, psFactory
            actionDroid_objLoc psDroid, DACTION_MOVE, psFactory, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y
        End If
    Case DORDER_GUARD
        aDroid(psDroid).order = psOrder
        If psOrder.psObj <> 0 Then
            objPos = SIMPLE_OBJECT_pos(psOrder.psObj)
            aDroid(psDroid).order.pos = Vector2i_(objPos.x, objPos.y)
        End If
        actionDroid psDroid, DACTION_NONE
    Case DORDER_RESTORE
        If Not electronicDroid(psDroid) Then
            GoTo switchEnd
        End If
        If SIMPLE_OBJECT_type(psOrder.psObj) <> OBJ_STRUCTURE Then
            ASSERT_FAILURE "false", "src/order.cpp:1913", "orderDroidBase", "orderDroidBase: invalid object type for Restore order"
            GoTo switchEnd
        End If
        aDroid(psDroid).order = psOrder
        objPos = SIMPLE_OBJECT_pos(psOrder.psObj)
        aDroid(psDroid).order.pos = Vector2i_(objPos.x, objPos.y)
        actionDroid_obj psDroid, DACTION_RESTORE, psOrder.psObj
    Case DORDER_REARM
        ' didn't get executed before
        If Not DROID_isVtolRearming(psDroid) Then
            aDroid(psDroid).order = psOrder
            actionDroid_obj psDroid, DACTION_MOVETOREARM, psOrder.psObj
            assignVTOLPad psDroid, psOrder.psObj
        End If
    Case DORDER_CIRCLE
        If Not DROID_isVtol(psDroid) Then
            GoTo switchEnd
        End If
        aDroid(psDroid).order = psOrder
        actionDroid_loc psDroid, DACTION_MOVE, psOrder.pos.x, psOrder.pos.y
    Case Else
        ASSERT_FAILURE "false", "src/order.cpp:1938", "orderDroidBase", "orderUnitBase: unknown order"
    End Select
switchEnd:

    syncDebugDroid_ "orderDroidBase", psDroid, Asc("+")
End Sub

' (VB) statsdef.h: getStatsName(_psStats) ((_psStats)->name.isEmpty() ? "" : (_psStats)->name.toUtf8().c_str()), for a STRUCTURE_STATS
Private Function getStatsName_(ByVal psStats As Long) As String
    getStatsName_ = asStructureStats(psStats).name_
End Function

' This function sends the droid an order. It uses sendDroidInfo() if mode == ModeQueue and orderDroidBase() if not.
Public Sub orderDroid(ByVal psDroid As Long, ByVal order As DroidOrderType, ByVal mode As QUEUE_MODE)
    Dim sOrder As DROID_ORDER_DATA

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/order.cpp:1949", "orderDroid", "orderUnit: Invalid unit pointer"
    If Not (order = DORDER_NONE Or _
            order = DORDER_RTR Or _
            order = DORDER_RTB Or _
            order = DORDER_RECYCLE Or _
            order = DORDER_TRANSPORTIN Or _
            order = DORDER_STOP Or _
            order = DORDER_HOLD) Then ASSERT_FAILURE "order == DORDER_NONE || order == DORDER_RTR || order == DORDER_RTB || order == DORDER_RECYCLE || order == DORDER_TRANSPORTIN || order == DORDER_STOP || order == DORDER_HOLD", "src/order.cpp:1951", "orderDroid", "orderUnit: Invalid order"

    sOrder = DroidOrder_(order)
    If mode = ModeQueue Then
        sendDroidInfo psDroid, sOrder, False
    Else
        orderClearDroidList psDroid
        orderDroidBase psDroid, sOrder
    End If
End Sub

' This function compares the current droid's order to the order.
' Returns true if they are the same, false else.
Public Function orderState(ByVal psDroid As Long, ByVal order As DroidOrderType) As Boolean
    If order = DORDER_RTR Then
        orderState = aDroid(psDroid).order.typ = DORDER_RTR Or aDroid(psDroid).order.typ = DORDER_RTR_SPECIFIED
        Exit Function
    End If

    orderState = aDroid(psDroid).order.typ = order
End Function

' This function returns true if the order is an acceptable order to give for a given location on the map.
Public Function validOrderForLoc(ByVal order As DroidOrderType) As Boolean
    validOrderForLoc = order = DORDER_NONE Or order = DORDER_MOVE Or order = DORDER_DISEMBARK Or _
                       order = DORDER_TRANSPORTOUT Or order = DORDER_TRANSPORTIN Or order = DORDER_TRANSPORTRETURN Or _
                       order = DORDER_GUARD Or order = DORDER_SCOUT Or order = DORDER_PATROL Or order = DORDER_CIRCLE
End Function

' This function sends the droid an order with a location.
' If the mode is ModeQueue, the order is added to the droid's order list using sendDroidInfo(), else, a DROID_ORDER_DATA is alloc, the old order list is erased, and the order is sent using orderDroidBase().
Public Sub orderDroidLoc(ByVal psDroid As Long, ByVal order As DroidOrderType, ByVal x As Long, ByVal y As Long, ByVal mode As QUEUE_MODE)
    Dim sOrder As DROID_ORDER_DATA

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/order.cpp:2001", "orderDroidLoc", "Invalid unit pointer": Exit Sub
    If Not (validOrderForLoc(order)) Then ASSERT_FAILURE "validOrderForLoc(order)", "src/order.cpp:2002", "orderDroidLoc", "Invalid order for location": Exit Sub

    sOrder = DroidOrder_pos(order, Vector2i_(x, y))
    If mode = ModeQueue Then
        sendDroidInfo psDroid, sOrder, False
        Exit Sub  ' Wait to receive our order before changing the droid.
    End If

    orderClearDroidList psDroid
    orderDroidBase psDroid, sOrder
End Sub

' This function attributes the order's location to (pX,pY) if the order is the same as the droid.
' Returns true if it was attributed and false if not.
Public Function orderStateLoc(ByVal psDroid As Long, ByVal order As DroidOrderType, pX As Long, pY As Long) As Boolean
    If order <> aDroid(psDroid).order.typ Then
        orderStateLoc = False
        Exit Function
    End If

    ' check the order is one with a location
    Select Case aDroid(psDroid).order.typ
    Case DORDER_MOVE
        pX = aDroid(psDroid).order.pos.x
        pY = aDroid(psDroid).order.pos.y
        orderStateLoc = True
        Exit Function
    Case Else
        ' not a location order - return false
    End Select

    orderStateLoc = False
End Function

' This function returns true if the order is a valid order to give to an object and false if it's not.
Public Function validOrderForObj(ByVal order As DroidOrderType) As Boolean
    validOrderForObj = order = DORDER_NONE Or order = DORDER_ATTACK Or order = DORDER_HELPBUILD Or _
                       order = DORDER_DEMOLISH Or order = DORDER_REPAIR Or order = DORDER_OBSERVE Or _
                       order = DORDER_FIRESUPPORT Or order = DORDER_RTR Or order = DORDER_EMBARK Or _
                       order = DORDER_ATTACKTARGET Or order = DORDER_COMMANDERSUPPORT Or order = DORDER_BUILDMODULE Or _
                       order = DORDER_GUARD Or order = DORDER_DROIDREPAIR Or order = DORDER_RESTORE Or _
                       order = DORDER_REARM Or order = DORDER_RECOVER Or order = DORDER_RTR_SPECIFIED
End Function

' This function sends an order with an object to the droid.
' If the mode is ModeQueue, the order is added to the droid's order list using sendDroidInfo(), else, a DROID_ORDER_DATA is alloc, the old order list is erased, and the order is sent using orderDroidBase().
Public Sub orderDroidObj(ByVal psDroid As Long, ByVal order As DroidOrderType, ByVal psObj As Long, ByVal mode As QUEUE_MODE)
    Dim sOrder As DROID_ORDER_DATA

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/order.cpp:2060", "orderDroidObj", "Invalid unit pointer"
    If Not (validOrderForObj(order)) Then ASSERT_FAILURE "validOrderForObj(order)", "src/order.cpp:2061", "orderDroidObj", "Invalid order for object"
    If Not (psObj <> 0) Then ASSERT_FAILURE "psObj != nullptr", "src/order.cpp:2062", "orderDroidObj", "Invalid object pointer": Exit Sub
    If Not (Not isBlueprint(psObj)) Then ASSERT_FAILURE "!isBlueprint(psObj)", "src/order.cpp:2063", "orderDroidObj", "Target %s is a blueprint", objInfo(psObj): Exit Sub
    If Not (SIMPLE_OBJECT_died(psObj) = 0) Then ASSERT_FAILURE "!psObj->died", "src/order.cpp:2064", "orderDroidObj", "Target dead": Exit Sub

    sOrder = DroidOrder_obj(order, psObj)
    If mode = ModeQueue Then 'ajl
        sendDroidInfo psDroid, sOrder, False
        Exit Sub  ' Wait for the order to be received before changing the droid.
    End If

    orderClearDroidList psDroid
    orderDroidBase psDroid, sOrder
End Sub

' This function returns the order's target if it has one, and NULL if the order is not a target order.
' @todo the first switch can be removed and substituted by orderState() function.
' @todo the use of this function is somewhat superfluous on some cases. Investigate.
Public Function orderStateObj(ByVal psDroid As Long, ByVal order As DroidOrderType) As Long
    Dim match As Boolean

    match = False

    Select Case order
    Case DORDER_BUILD, DORDER_LINEBUILD, DORDER_HELPBUILD
        If aDroid(psDroid).order.typ = DORDER_BUILD Or _
           aDroid(psDroid).order.typ = DORDER_HELPBUILD Or _
           aDroid(psDroid).order.typ = DORDER_LINEBUILD Then
            match = True
        End If
    Case DORDER_ATTACK, DORDER_FIRESUPPORT, DORDER_OBSERVE, DORDER_DEMOLISH, DORDER_DROIDREPAIR, DORDER_REARM, DORDER_GUARD
        If aDroid(psDroid).order.typ = order Then
            match = True
        End If
    Case DORDER_RTR
        If aDroid(psDroid).order.typ = DORDER_RTR Or _
           aDroid(psDroid).order.typ = DORDER_RTR_SPECIFIED Then
            match = True
        End If
    Case Else
    End Select

    If Not match Then
        orderStateObj = 0
        Exit Function
    End If

    ' check the order is one with an object
    Select Case aDroid(psDroid).order.typ
    Case DORDER_BUILD, DORDER_LINEBUILD
        If aDroid(psDroid).action = DACTION_BUILD Or _
           aDroid(psDroid).action = DACTION_BUILDWANDER Then
            orderStateObj = aDroid(psDroid).order.psObj
            Exit Function
        End If
    Case DORDER_HELPBUILD
        If aDroid(psDroid).action = DACTION_BUILD Or _
           aDroid(psDroid).action = DACTION_BUILDWANDER Or _
           aDroid(psDroid).action = DACTION_MOVETOBUILD Then
            orderStateObj = aDroid(psDroid).order.psObj
            Exit Function
        End If
    'case DORDER_HELPBUILD:
    Case DORDER_ATTACK, DORDER_FIRESUPPORT, DORDER_OBSERVE, DORDER_DEMOLISH, DORDER_RTR, DORDER_RTR_SPECIFIED, _
         DORDER_DROIDREPAIR, DORDER_REARM, DORDER_GUARD
        orderStateObj = aDroid(psDroid).order.psObj
        Exit Function
    Case Else
        ' not an object order - return false
        orderStateObj = 0
        Exit Function
    End Select

    orderStateObj = 0
End Function

' This function sends the droid an order with a location and stats.
' If the mode is ModeQueue, the order is added to the droid's order list using sendDroidInfo(), else, a DROID_ORDER_DATA is alloc, the old order list is erased, and the order is sent using orderDroidBase().
Public Sub orderDroidStatsLocDir(ByVal psDroid As Long, ByVal order As DroidOrderType, ByVal psStats As Long, ByVal x As Long, ByVal y As Long, ByVal direction As Long, ByVal mode As QUEUE_MODE)
    Dim sOrder As DROID_ORDER_DATA

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/order.cpp:2171", "orderDroidStatsLocDir", "Invalid unit pointer"
    If Not (order = DORDER_BUILD) Then ASSERT_FAILURE "order == DORDER_BUILD", "src/order.cpp:2172", "orderDroidStatsLocDir", "Invalid order for location"

    sOrder = DroidOrder_stats(order, psStats, Vector2i_(x, y), direction)
    If mode = ModeQueue Then
        sendDroidInfo psDroid, sOrder, False
        Exit Sub  ' Wait for our order before changing the droid.
    End If

    orderClearDroidList psDroid
    orderDroidBase psDroid, sOrder
End Sub

' This function adds that order to the droid's list using sendDroidInfo().
' @todo seems closely related with orderDroidStatsLocDir(). See if this one can be incorporated on it.
Public Sub orderDroidStatsLocDirAdd(ByVal psDroid As Long, ByVal order As DroidOrderType, ByVal psStats As Long, ByVal x As Long, ByVal y As Long, ByVal direction As Long, Optional ByVal add As Boolean = True)
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/order.cpp:2191", "orderDroidStatsLocDirAdd", "Invalid unit pointer"

    ' can only queue build orders with this function
    If order <> DORDER_BUILD Then
        Exit Sub
    End If

    sendDroidInfo psDroid, DroidOrder_stats(order, psStats, Vector2i_(x, y), direction), add
End Sub

' Equivalent to orderDroidStatsLocDir(), but uses two locations.
Public Sub orderDroidStatsTwoLocDir(ByVal psDroid As Long, ByVal order As DroidOrderType, ByVal psStats As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal direction As Long, ByVal mode As QUEUE_MODE)
    Dim sOrder As DROID_ORDER_DATA

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/order.cpp:2206", "orderDroidStatsTwoLocDir", "Invalid unit pointer"
    If Not (order = DORDER_LINEBUILD) Then ASSERT_FAILURE "order == DORDER_LINEBUILD", "src/order.cpp:2207", "orderDroidStatsTwoLocDir", "Invalid order for location"

    sOrder = DroidOrder_stats2(order, psStats, Vector2i_(x1, y1), Vector2i_(x2, y2), direction)
    If mode = ModeQueue Then
        sendDroidInfo psDroid, sOrder, False
        Exit Sub  ' Wait for our order before changing the droid.
    End If

    orderClearDroidList psDroid
    orderDroidBase psDroid, sOrder
End Sub

' Equivalent to orderDroidStatsLocDirAdd(), but uses two locations.
' @todo seems closely related with orderDroidStatsTwoLocDir(). See if this can be incorporated on it.
Public Sub orderDroidStatsTwoLocDirAdd(ByVal psDroid As Long, ByVal order As DroidOrderType, ByVal psStats As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal direction As Long)
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/order.cpp:2226", "orderDroidStatsTwoLocDirAdd", "Invalid unit pointer"
    If Not (order = DORDER_LINEBUILD) Then ASSERT_FAILURE "order == DORDER_LINEBUILD", "src/order.cpp:2227", "orderDroidStatsTwoLocDirAdd", "Invalid order for location"

    sendDroidInfo psDroid, DroidOrder_stats2(order, psStats, Vector2i_(x1, y1), Vector2i_(x2, y2), direction), True
End Sub

' This function returns false if droid's order and order don't match or the order is not a location order. Else ppsStats = psDroid->psTarStats, (pX,pY) = psDroid.(orderX,orderY) and it returns true.
' @todo seems closely related to orderStateLoc()
Public Function orderStateStatsLoc(ByVal psDroid As Long, ByVal order As DroidOrderType, ppsStats As Long) As Boolean
    Dim match As Boolean

    match = False

    Select Case order
    Case DORDER_BUILD, DORDER_LINEBUILD
        If aDroid(psDroid).order.typ = DORDER_BUILD Or _
           aDroid(psDroid).order.typ = DORDER_LINEBUILD Then
            match = True
        End If
    Case Else
    End Select
    If Not match Then
        orderStateStatsLoc = False
        Exit Function
    End If

    ' check the order is one with stats and a location
    Select Case aDroid(psDroid).order.typ
    Case DORDER_BUILD, DORDER_LINEBUILD
        If aDroid(psDroid).action = DACTION_MOVETOBUILD Then
            ppsStats = aDroid(psDroid).order.psStats
            orderStateStatsLoc = True
            Exit Function
        End If
    Case Else
        ' not a stats/location order - return false
        orderStateStatsLoc = False
        Exit Function
    End Select

    orderStateStatsLoc = False
End Function

' (VB) psDroid->asOrderList.resize(n) (or push_back's growth): the array keeps at least n elements
Private Sub orderListResize(ByVal psDroid As Long, ByVal n As Long)
    Dim i As Long, oldSize As Long
    oldSize = aDroid(psDroid).asOrderList_size
    If n > oldSize Then
        If oldSize = 0 Then
            ReDim aDroid(psDroid).asOrderList(0 To n - 1)
        ElseIf UBound(aDroid(psDroid).asOrderList) < n - 1 Then
            ReDim Preserve aDroid(psDroid).asOrderList(0 To n - 1)
        End If
        For i = oldSize To n - 1
            DroidOrder_construct aDroid(psDroid).asOrderList(i)
        Next
    End If
    aDroid(psDroid).asOrderList_size = n
End Sub

' @todo needs documentation.
Public Sub orderDroidAddPending(ByVal psDroid As Long, psOrder As DROID_ORDER_DATA)
    Dim position As Vector3i, objPos As Vector3i, imd As Long

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/order.cpp:2282", "orderDroidAddPending", "Invalid unit pointer": Exit Sub

    orderListResize psDroid, aDroid(psDroid).asOrderList_size + 1
    aDroid(psDroid).asOrderList(aDroid(psDroid).asOrderList_size - 1) = psOrder

    ' Only display one arrow, bOrderEffectDisplayed must be set to false once per arrow.
    If Not bOrderEffectDisplayed Then
        position = Vector3i_(0, 0, 0)
        If psOrder.psObj = 0 Then
            position.x = psOrder.pos.x
            position.z = psOrder.pos.y
        Else
            objPos = SIMPLE_OBJECT_pos(psOrder.psObj)
            position = Vector3i_(objPos.x, objPos.z, objPos.y)  ' pos.xzy()
        End If
        position.y = map_Height(position.x, position.z) + 32
        If psOrder.psObj <> 0 Then
            imd = BASE_OBJECT_sDisplay_imd(psOrder.psObj)
            If imd <> 0 Then
                position.y = position.y + imdMaxY(imd)
            End If
        End If
        addEffect position, EFFECT_WAYPOINT, WAYPOINT_TYPE, False, 0, 0, 0
        bOrderEffectDisplayed = True
    End If
End Sub

' Add an order to a droid's order list
' @todo needs better documentation.
Public Sub orderDroidAdd(ByVal psDroid As Long, psOrder As DROID_ORDER_DATA)
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/order.cpp:2315", "orderDroidAdd", "Invalid unit pointer": Exit Sub

    If Not ULT(aDroid(psDroid).listSize, aDroid(psDroid).asOrderList_size) Then
        ' Make more room to store the order.
        orderListResize psDroid, aDroid(psDroid).asOrderList_size + 1
    End If

    aDroid(psDroid).asOrderList(aDroid(psDroid).listSize) = psOrder
    aDroid(psDroid).listSize = aDroid(psDroid).listSize + 1

    ' if not doing anything - do it immediately
    If aDroid(psDroid).listSize <= 1 And _
       (aDroid(psDroid).order.typ = DORDER_NONE Or _
        aDroid(psDroid).order.typ = DORDER_GUARD Or _
        aDroid(psDroid).order.typ = DORDER_PATROL Or _
        aDroid(psDroid).order.typ = DORDER_CIRCLE Or _
        aDroid(psDroid).order.typ = DORDER_HOLD) Then
        orderDroidList psDroid
    End If
End Sub

' This function goes to the droid's order list and sets a new order to it from its order list.
Public Function orderDroidList(ByVal psDroid As Long) As Boolean
    Dim sOrder As DROID_ORDER_DATA

    If aDroid(psDroid).listSize > 0 Then
        ' there are some orders to give
        sOrder = aDroid(psDroid).asOrderList(0)
        orderDroidListEraseRange psDroid, 0, 1

        Select Case sOrder.typ
        Case DORDER_MOVE, DORDER_SCOUT, DORDER_DISEMBARK
            If Not (sOrder.psObj = 0 And sOrder.psStats = -1) Then ASSERT_FAILURE "sOrder.psObj == nullptr && sOrder.psStats == nullptr", "src/order.cpp:2353", "orderDroidList", "Extra %s parameters.", getDroidOrderName(sOrder.typ)
            sOrder.psObj = 0
            sOrder.psStats = -1
        Case DORDER_ATTACK, DORDER_REPAIR, DORDER_OBSERVE, DORDER_DROIDREPAIR, DORDER_FIRESUPPORT, DORDER_DEMOLISH, _
             DORDER_HELPBUILD, DORDER_BUILDMODULE, DORDER_RECOVER
            If Not (sOrder.psStats = -1) Then ASSERT_FAILURE "sOrder.psStats == nullptr", "src/order.cpp:2366", "orderDroidList", "Extra %s parameters.", getDroidOrderName(sOrder.typ)
            sOrder.psStats = -1
        Case DORDER_BUILD, DORDER_LINEBUILD
            If Not (sOrder.psObj = 0) Then ASSERT_FAILURE "sOrder.psObj == nullptr", "src/order.cpp:2371", "orderDroidList", "Extra %s parameters.", getDroidOrderName(sOrder.typ)
            sOrder.psObj = 0
        Case Else
            ASSERT_FAILURE "false", "src/order.cpp:2375", "orderDroidList", "orderDroidList: Invalid order"
            orderDroidList = False
            Exit Function
        End Select

        orderDroidBase psDroid, sOrder
        syncDebugDroid_ "orderDroidList", psDroid, Asc("o")

        orderDroidList = True
        Exit Function
    End If

    orderDroidList = False
End Function

' This function goes to the droid's order list and erases its elements from indexBegin to indexEnd.
Public Sub orderDroidListEraseRange(ByVal psDroid As Long, ByVal indexBegin As Long, ByVal indexEnd As Long)
    Dim i As Long, n As Long

    ' Erase elements
    indexEnd = MIN_(indexEnd, aDroid(psDroid).asOrderList_size)  ' Do nothing if trying to pop an empty list.
    n = indexEnd - indexBegin
    If n > 0 Then
        For i = indexEnd To aDroid(psDroid).asOrderList_size - 1
            aDroid(psDroid).asOrderList(i - n) = aDroid(psDroid).asOrderList(i)
        Next
        aDroid(psDroid).asOrderList_size = aDroid(psDroid).asOrderList_size - n
    End If

    ' Update indices into list.
    aDroid(psDroid).listSize = aDroid(psDroid).listSize - (MIN_(indexEnd, aDroid(psDroid).listSize) - MIN_(indexBegin, aDroid(psDroid).listSize))
    aDroid(psDroid).listPendingBegin = aDroid(psDroid).listPendingBegin - (MIN_(indexEnd, aDroid(psDroid).listPendingBegin) - MIN_(indexBegin, aDroid(psDroid).listPendingBegin))
End Sub

' This function clears all the synchronised orders from the list, calling orderDroidListEraseRange() from 0 to psDroid->listSize.
Public Sub orderClearDroidList(ByVal psDroid As Long)
    syncDebug "orderClearDroidList", "droid%d list cleared", aDroid(psDroid).id
    orderDroidListEraseRange psDroid, 0, aDroid(psDroid).listSize
End Sub

' This function clears all the orders from droid's order list that don't have target as psTarget.
Public Sub orderClearTargetFromDroidList(ByVal psDroid As Long, ByVal psTarget As Long)
    Dim i As Long
    i = 0
    Do While i < aDroid(psDroid).asOrderList_size
        If aDroid(psDroid).asOrderList(i).psObj = psTarget Then
            If i < aDroid(psDroid).listSize Then
                syncDebug "orderClearTargetFromDroidList", "droid%d list erase%d", aDroid(psDroid).id, SIMPLE_OBJECT_id(psTarget)
            End If
            orderDroidListEraseRange psDroid, i, i + 1
            i = i - 1  ' If this underflows, the ++i will overflow it back.
        End If
        i = i + 1
    Loop
End Sub

' This function checks its order list: if a given order needs a target and the target has died, the order is removed from the list.
Public Sub orderCheckList(ByVal psDroid As Long)
    Dim i As Long, psTarget As Long
    i = 0
    Do While i < aDroid(psDroid).asOrderList_size
        psTarget = aDroid(psDroid).asOrderList(i).psObj
        If psTarget <> 0 Then
            If SIMPLE_OBJECT_died(psTarget) <> 0 Then
                If i < aDroid(psDroid).listSize Then
                    syncDebugObject_ "orderCheckList", psTarget, Asc("-")
                    syncDebug "orderCheckList", "droid%d list erase dead droid%d", aDroid(psDroid).id, SIMPLE_OBJECT_id(psTarget)
                End If
                orderDroidListEraseRange psDroid, i, i + 1
                i = i - 1  ' If this underflows, the ++i will overflow it back.
            End If
        End If
        i = i + 1
    Loop
End Sub

' This function sends the droid an order with a location using sendDroidInfo().
' @todo it is very close to what orderDroidLoc() function does. Suggestion to refract them.
Private Function orderDroidLocAdd(ByVal psDroid As Long, ByVal order As DroidOrderType, ByVal x As Long, ByVal y As Long, Optional ByVal add As Boolean = True) As Boolean
    ' can only queue move, scout, and disembark orders
    If order <> DORDER_MOVE And order <> DORDER_SCOUT And order <> DORDER_DISEMBARK Then
        orderDroidLocAdd = False
        Exit Function
    End If

    sendDroidInfo psDroid, DroidOrder_pos(order, Vector2i_(x, y)), add

    orderDroidLocAdd = True
End Function

' This function sends the droid an order with a location using sendDroidInfo().
' @todo it is very close to what orderDroidObj() function does. Suggestion to refract them.
Private Function orderDroidObjAdd(ByVal psDroid As Long, order As DROID_ORDER_DATA, ByVal add As Boolean) As Boolean
    If Not (Not isBlueprint(order.psObj)) Then ASSERT_FAILURE "!isBlueprint(order.psObj)", "src/order.cpp:2470", "orderDroidObjAdd", "Target %s for queue is a blueprint", objInfo(order.psObj)

    ' check can queue the order
    Select Case order.typ
    Case DORDER_ATTACK, DORDER_REPAIR, DORDER_OBSERVE, DORDER_DROIDREPAIR, DORDER_FIRESUPPORT, DORDER_DEMOLISH, _
         DORDER_HELPBUILD, DORDER_BUILDMODULE
    Case Else
        orderDroidObjAdd = False
        Exit Function
    End Select

    sendDroidInfo psDroid, order, add

    orderDroidObjAdd = True
End Function

' This function returns an order which is assigned according to the location and droid. Uses altOrder flag to choose between a direct order or an altOrder.
Public Function chooseOrderLoc(ByVal psDroid As Long, ByVal x As Long, ByVal y As Long, ByVal altOrder As Boolean) As DroidOrderType
    Dim order As DroidOrderType, propulsion As Long

    order = DORDER_NONE
    propulsion = asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType

    If DROID_isTransporter(psDroid) And game.typ = LEVEL_TYPE_CAMPAIGN Then
        ' transports can't be controlled in campaign
        chooseOrderLoc = DORDER_NONE
        Exit Function
    End If

    ' default to move; however, we can only end up on a tile
    ' where can stay, ie VTOLs must be able to land as well
    If DROID_isVtol(psDroid) Then
        propulsion = PROPULSION_TYPE_WHEELED
    End If
    If Not fpathBlockingTile(map_coord(x), map_coord(y), propulsion) Then
        order = DORDER_MOVE
    End If

    ' scout if alt was pressed
    If altOrder Then
        order = DORDER_SCOUT
        If DROID_isVtol(psDroid) Then
            ' Patrol if in a VTOL
            order = DORDER_PATROL
        End If
    End If

    ' and now we want Transporters to fly! - in multiPlayer!!
    If DROID_isTransporter(psDroid) And game.typ = LEVEL_TYPE_SKIRMISH Then
        ' in MultiPlayer - if ALT-key is pressed then need to get the Transporter
        ' to fly to location and all units disembark
        If altOrder Then
            order = DORDER_DISEMBARK
        End If
    ElseIf secondaryGetState(psDroid, DSO_CIRCLE, ModeQueue) = DSS_CIRCLE_SET Then  ' ModeQueue here means to check whether we pressed the circle button, whether or not it synched yet. The reason for this weirdness is that a circle order makes no sense as a secondary state in the first place (the circle button _should_ have been only in the UI, not in the game state..!), so anything dealing with circle orders will necessarily be weird.
        order = DORDER_CIRCLE
        secondarySetState psDroid, DSO_CIRCLE, DSS_NONE
    ElseIf secondaryGetState(psDroid, DSO_PATROL, ModeQueue) = DSS_PATROL_SET Then  ' ModeQueue here means to check whether we pressed the patrol button, whether or not it synched yet. The reason for this weirdness is that a patrol order makes no sense as a secondary state in the first place (the patrol button _should_ have been only in the UI, not in the game state..!), so anything dealing with patrol orders will necessarily be weird.
        order = DORDER_PATROL
        secondarySetState psDroid, DSO_PATROL, DSS_NONE
    End If

    chooseOrderLoc = order
End Function

' This function sends the selected droids an order to given a location. If a delivery point is selected, it is moved to a new location.
' If add is true then the order is queued.
' This function should only be called from UI.
Public Sub orderSelectedLoc(ByVal player As Long, ByVal x As Long, ByVal y As Long, ByVal add As Boolean)
    Dim order As DroidOrderType, it As Long, psCurr As Long, c As Boolean

    'if were in build select mode ignore all other clicking
    If intBuildSelectMode() Then
        Exit Sub
    End If

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/order.cpp:2567", "orderSelectedLoc", "Invalid player: %u", player: Exit Sub

    ' note that an order list graphic needs to be displayed
    bOrderEffectDisplayed = False

    it = apsDroidLists(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aDroid(psCurr).selected <> 0 Then
            ' can't use bMultiPlayer since multimsg could be off.
            If aDroid(psCurr).droidType = DROID_SUPERTRANSPORTER And game.typ = LEVEL_TYPE_CAMPAIGN Then
                ' Transport in campaign cannot be controlled by players
                DeSelectDroid psCurr
                GoTo continue_
            End If

            order = chooseOrderLoc(psCurr, x, y, specialOrderKeyDown())
            ' see if the order can be added to the list
            If order <> DORDER_NONE Then
                c = False
                If add Then c = orderDroidLocAdd(psCurr, order, x, y)
                If Not c Then
                    ' if not just do it straight off
                    orderDroidLoc psCurr, order, x, y, ModeQueue
                End If
            End If
        End If
continue_:
        it = lnNext(it)
    Loop
End Sub

Private Function highestQueuedModule_order(order As DROID_ORDER_DATA, ByVal structure_ As Long, ByVal prevHighestQueuedModule As Long) As Long
    Dim thisQueuedModule As Long, orderStructure As Long
    thisQueuedModule = -1
    Select Case order.typ
    Case DORDER_BUILDMODULE
        If order.psObj = structure_ Then  ' Order must be for this structure.
            thisQueuedModule = order.index  ' Order says which module to build.
        End If
    Case DORDER_BUILD, DORDER_HELPBUILD
        ' Current order is weird, the DORDER_BUILDMODULE mutates into a DORDER_BUILD, and we use the order.pos instead of order.psObj.
        ' Also, might be DORDER_BUILD if selecting the module from the menu before clicking on the structure.
        orderStructure = castStructure(psMapTiles(worldTile(order.pos.x, order.pos.y)).psObject)
        If orderStructure = structure_ Then  ' Order must be for this structure.
            If order.psStats = aStruct(orderStructure).pStructureType Or order.psStats = getModuleStat(orderStructure) Then
                thisQueuedModule = nextModuleToBuild(structure_, prevHighestQueuedModule)
            End If
        End If
    Case Else
    End Select
    highestQueuedModule_order = MAX_(prevHighestQueuedModule, thisQueuedModule)
End Function

Private Function highestQueuedModule(ByVal droid_ As Long, ByVal structure_ As Long) As Long
    Dim module As Long, n As Long
    module = highestQueuedModule_order(aDroid(droid_).order, structure_, -1)
    For n = aDroid(droid_).listPendingBegin To aDroid(droid_).asOrderList_size - 1
        module = highestQueuedModule_order(aDroid(droid_).asOrderList(n), structure_, module)
    Next
    highestQueuedModule = module
End Function

' This function returns an order according to the droid, object (target) and altOrder.
Public Function chooseOrderObj(ByVal psDroid As Long, ByVal psObj As Long, ByVal altOrder As Boolean) As DROID_ORDER_DATA
    Dim order As DROID_ORDER_DATA, psStruct As Long, i As Long, sensorType As Long, moduleIndex As Long
    Dim ObjDepartedAlly As Boolean, objType As OBJECT_TYPE, objPlayer As Long, c As Boolean

    order = DroidOrder_(DORDER_NONE)
    objType = SIMPLE_OBJECT_type(psObj)
    objPlayer = SIMPLE_OBJECT_player(psObj)

    If DROID_isTransporter(psDroid) Then
        'in multiPlayer, need to be able to get Transporter repaired
        If bMultiPlayer Then
            If aiCheckAlliances(objPlayer, aDroid(psDroid).player) And _
               objType = OBJ_STRUCTURE Then
                psStruct = psObj
                If Not (psObj <> 0) Then ASSERT_FAILURE "psObj != nullptr", "src/order.cpp:2649", "chooseOrderObj", "Invalid structure pointer": chooseOrderObj = DroidOrder_(DORDER_NONE): Exit Function
                If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REPAIR_FACILITY And _
                   aStruct(psStruct).status = SS_BUILT Then
                    chooseOrderObj = DroidOrder_obj(DORDER_RTR_SPECIFIED, psObj)
                    Exit Function
                End If
            End If
        End If
        chooseOrderObj = DroidOrder_(DORDER_NONE)
        Exit Function
    End If

    If altOrder And (objType = OBJ_DROID Or objType = OBJ_STRUCTURE) And aDroid(psDroid).player = objPlayer Then
        If aDroid(psDroid).droidType = DROID_SENSOR Then
            chooseOrderObj = DroidOrder_obj(DORDER_OBSERVE, psObj)
            Exit Function
        ElseIf (aDroid(psDroid).droidType = DROID_REPAIR Or _
                aDroid(psDroid).droidType = DROID_CYBORG_REPAIR) And objType = OBJ_DROID Then
            chooseOrderObj = DroidOrder_obj(DORDER_DROIDREPAIR, psObj)
            Exit Function
        ElseIf (aDroid(psDroid).droidType = DROID_WEAPON) Or DROID_isCyborg(psDroid) Or _
               (aDroid(psDroid).droidType = DROID_COMMAND) Then
            chooseOrderObj = DroidOrder_obj(DORDER_ATTACK, psObj)
            Exit Function
        End If
    End If
    'check for transporters first
    c = False
    If objType = OBJ_DROID Then
        If DROID_isTransporter(psObj) Then c = objPlayer = aDroid(psDroid).player
    End If
    If c Then
        order = DroidOrder_obj(DORDER_EMBARK, psObj)
    ' go to recover an artifact/oil drum - don't allow VTOL's to get this order
    ElseIf isArtefactOrOilDrum(psObj) Then
        If Not DROID_isVtol(psDroid) Then
            order = DroidOrder_obj(DORDER_RECOVER, psObj)
        End If
    ' else default to attack if the droid has a weapon
    ElseIf aDroid(psDroid).numWeaps > 0 _
           And objPlayer <> aDroid(psDroid).player _
           And Not aiCheckAlliances(objPlayer, aDroid(psDroid).player) Then
        ' check valid weapon/prop combination
        For i = 0 To MAX_WEAPONS - 1
            If validTarget(psDroid, psObj, i) Then
                order = DroidOrder_obj(DORDER_ATTACK, psObj)
                Exit For
            End If
        Next
    ElseIf aDroid(psDroid).droidType = DROID_SENSOR _
           And objPlayer <> aDroid(psDroid).player _
           And Not aiCheckAlliances(objPlayer, aDroid(psDroid).player) Then
        sensorType = asSensorStats(DROID_getSensorStats(psDroid)).typ
        'check for standard sensor or VTOL intercept sensor
        If sensorType = STANDARD_SENSOR _
           Or sensorType = VTOL_INTERCEPT_SENSOR _
           Or sensorType = SUPER_SENSOR Then
            ' a sensor droid observing an object
            order = DroidOrder_obj(DORDER_OBSERVE, psObj)
        End If
    ElseIf droidSensorDroidWeapon(psObj, psDroid) Then
        ' got an indirect weapon droid or vtol doing fire support
        order = DroidOrder_obj(DORDER_FIRESUPPORT, psObj)
        setSensorAssigned
    ElseIf joinsCommander(psDroid, psObj) Then
        ' get a droid to join a command droids group
        DeSelectDroid psDroid
        order = DroidOrder_obj(DORDER_COMMANDERSUPPORT, psObj)
    'repair droid
    ElseIf repairsDroid(psDroid, psObj) Then
        order = DroidOrder_obj(DORDER_DROIDREPAIR, psObj)
    ' guarding constructor droids
    ElseIf guardsDroid(psDroid, psObj) Then
        order = DroidOrder_obj(DORDER_GUARD, psObj)
        assignSensorTarget psObj
        aDroid(psDroid).selected = 0
    ElseIf aiCheckAlliances(objPlayer, aDroid(psDroid).player) And _
           objType = OBJ_STRUCTURE Then
        psStruct = psObj
        If Not (psObj <> 0) Then ASSERT_FAILURE "psObj != nullptr", "src/order.cpp:2767", "chooseOrderObj", "Invalid structure pointer": chooseOrderObj = DroidOrder_(DORDER_NONE): Exit Function

        ' check whether construction droid
        If aDroid(psDroid).droidType = DROID_CONSTRUCT Or _
           aDroid(psDroid).droidType = DROID_CYBORG_CONSTRUCT Then
            If ctrlShiftDown() Then
                moduleIndex = nextModuleToBuild(psStruct, highestQueuedModule(psDroid, psStruct))
            Else
                moduleIndex = nextModuleToBuild(psStruct, -1)
            End If

            'Re-written to allow demolish order to be added to the queuing system
            ObjDepartedAlly = False
            If bMultiPlayer Then
                If NetPlay_players_difficulty(objPlayer) = AIDifficulty_HUMAN Then
                    ObjDepartedAlly = Not NetPlay_players_allocated(objPlayer)
                End If
            End If
            c = False
            If intDemolishSelectMode() Then
                c = (objPlayer = aDroid(psDroid).player Or ObjDepartedAlly)
            End If
            If c Then
                'check to see if anything is currently trying to build the structure
                'can't build and demolish at the same time!
                If aStruct(psStruct).status = SS_BUILT Then
                    order = DroidOrder_obj(DORDER_DEMOLISH, psObj)
                ElseIf Not checkDroidsBuilding(psStruct) Then
                    order = DroidOrder_obj(DORDER_DEMOLISH, psObj)
                End If
            'check for non complete structures
            ElseIf aStruct(psStruct).status <> SS_BUILT Then
                'if something else is demolishing, then help demolish
                If checkDroidsDemolishing(psStruct) Then
                    order = DroidOrder_obj(DORDER_DEMOLISH, psObj)
                'else help build
                Else
                    order = DroidOrder_obj(DORDER_HELPBUILD, psObj)
                    If moduleIndex > 0 Then
                        order = DroidOrder_objIndex(DORDER_BUILDMODULE, psObj, moduleIndex)  ' Try scheduling a module, instead.
                    End If
                End If
            ElseIf ULT(aStruct(psStruct).body, STRUCTURE_structureBody(psStruct)) Then
                order = DroidOrder_obj(DORDER_REPAIR, psObj)
            'check if can build a module
            ElseIf moduleIndex > 0 Then
                order = DroidOrder_objIndex(DORDER_BUILDMODULE, psObj, moduleIndex)
            End If
        End If

        If order.typ = DORDER_NONE Then
            ' check repair facility and in need of repair
            c = False
            If electronicDroid(psDroid) Then
                c = aStruct(psStruct).resistance < structureResistance(aStruct(psStruct).pStructureType, aStruct(psStruct).player)
            End If
            If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REPAIR_FACILITY And _
               aStruct(psStruct).status = SS_BUILT Then
                order = DroidOrder_obj(DORDER_RTR_SPECIFIED, psObj)
            ElseIf c Then
                'psStruct->resistance < (SDWORD)(psStruct->pStructureType->resistance))
                order = DroidOrder_obj(DORDER_RESTORE, psObj)
            'check for counter battery assignment
            ElseIf structSensorDroidWeapon(psStruct, psDroid) Then
                order = DroidOrder_obj(DORDER_FIRESUPPORT, psObj)
                'inform display system
                setSensorAssigned
                'deselect droid
                DeSelectDroid psDroid
            'REARM VTOLS
            ElseIf DROID_isVtol(psDroid) Then
                'default to no order
                'check if rearm pad
                If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REARM_PAD Then
                    'don't bother checking cos we want it to go there if directed
                    order = DroidOrder_obj(DORDER_REARM, psObj)
                End If
            ' Some droids shouldn't be guarding
            ElseIf (aDroid(psDroid).droidType = DROID_WEAPON Or _
                    aDroid(psDroid).droidType = DROID_CYBORG Or _
                    aDroid(psDroid).droidType = DROID_CYBORG_SUPER) _
                   And proj_Direct(DROID_getWeaponStats(psDroid, 0)) Then
                order = DroidOrder_obj(DORDER_GUARD, psObj)
            End If
        End If
    End If

    chooseOrderObj = order
End Function

' (VB) chooseOrderObj: psObj->type == OBJ_FEATURE && (((FEATURE *)psObj)->psStats->subType == FEAT_GEN_ARTE ||
' ((FEATURE *)psObj)->psStats->subType == FEAT_OIL_DRUM)
Private Function isArtefactOrOilDrum(ByVal psObj As Long) As Boolean
    If SIMPLE_OBJECT_type(psObj) = OBJ_FEATURE Then
        isArtefactOrOilDrum = asFeatureStats(aFeature(psObj).psStats).subType = FEAT_GEN_ARTE Or _
                              asFeatureStats(aFeature(psObj).psStats).subType = FEAT_OIL_DRUM
    End If
End Function

' (VB) chooseOrderObj: psObj->player == psDroid->player && psObj->type == OBJ_DROID && ((DROID *)psObj)->droidType == DROID_COMMAND &&
' psDroid->droidType != DROID_COMMAND && psDroid->droidType != DROID_CONSTRUCT && psDroid->droidType != DROID_CYBORG_CONSTRUCT
Private Function joinsCommander(ByVal psDroid As Long, ByVal psObj As Long) As Boolean
    If SIMPLE_OBJECT_player(psObj) = aDroid(psDroid).player And SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        joinsCommander = aDroid(psObj).droidType = DROID_COMMAND And _
                         aDroid(psDroid).droidType <> DROID_COMMAND And _
                         aDroid(psDroid).droidType <> DROID_CONSTRUCT And _
                         aDroid(psDroid).droidType <> DROID_CYBORG_CONSTRUCT
    End If
End Function

' (VB) chooseOrderObj: aiCheckAlliances(psObj->player, psDroid->player) && psObj->type == OBJ_DROID &&
' (psDroid->droidType == DROID_REPAIR || psDroid->droidType == DROID_CYBORG_REPAIR) && ((DROID *)psObj)->isDamaged()
Private Function repairsDroid(ByVal psDroid As Long, ByVal psObj As Long) As Boolean
    If aiCheckAlliances(SIMPLE_OBJECT_player(psObj), aDroid(psDroid).player) Then
        If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
            If aDroid(psDroid).droidType = DROID_REPAIR Or aDroid(psDroid).droidType = DROID_CYBORG_REPAIR Then
                repairsDroid = DROID_isDamaged(psObj)
            End If
        End If
    End If
End Function

' (VB) chooseOrderObj: aiCheckAlliances(psObj->player, psDroid->player) && psObj->type == OBJ_DROID &&
' (target constructor, sensor or allied commander) && (psDroid weapon, cyborg or super cyborg) && proj_Direct(psDroid->getWeaponStats(0))
Private Function guardsDroid(ByVal psDroid As Long, ByVal psObj As Long) As Boolean
    Dim t As DROID_TYPE
    If aiCheckAlliances(SIMPLE_OBJECT_player(psObj), aDroid(psDroid).player) Then
        If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
            t = aDroid(psObj).droidType
            If t = DROID_CONSTRUCT Or t = DROID_CYBORG_CONSTRUCT Or t = DROID_SENSOR Or _
               (t = DROID_COMMAND And SIMPLE_OBJECT_player(psObj) <> aDroid(psDroid).player) Then
                If aDroid(psDroid).droidType = DROID_WEAPON Or _
                   aDroid(psDroid).droidType = DROID_CYBORG Or _
                   aDroid(psDroid).droidType = DROID_CYBORG_SUPER Then
                    guardsDroid = proj_Direct(DROID_getWeaponStats(psDroid, 0))
                End If
            End If
        End If
    End If
End Function

' This function runs through all the player's droids and if the droid is vtol and is selected and is attacking, uses audio_QueueTrack() to play a sound.
' @todo this function has variable psObj unused. Consider removing it from argument.
' @todo this function runs through all the player's droids, but only uses the selected ones. Consider an efficiency improvement in here.
' @todo current scope of this function is quite small. Consider refactoring it.
Private Sub orderPlayOrderObjAudio(ByVal player As Long, ByVal psObj As Long)
    Dim it As Long, psDroid As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/order.cpp:2873", "orderPlayOrderObjAudio", "Invalid player: %u", player: Exit Sub

    ' loop over selected droids
    it = apsDroidLists(player).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected <> 0 Then
            ' currently only looks for VTOL
            If DROID_isVtol(psDroid) Then
                Select Case aDroid(psDroid).order.typ
                Case DORDER_ATTACK
                    audio_QueueTrack ID_SOUND_ON_OUR_WAY2
                Case Else
                End Select
            End If

            ' only play audio once
            Exit Do
        End If
        it = lnNext(it)
    Loop
End Sub

' This function sends orders to all the selected droids according to the object.
' If add is true, the orders are queued.
' @todo this function runs through all the player's droids, but only uses the selected ones. Consider an efficiency improvement in here.
Public Sub orderSelectedObjAdd(ByVal player As Long, ByVal psObj As Long, ByVal add As Boolean)
    Dim it As Long, psCurr As Long, order As DROID_ORDER_DATA, objPos As Vector3i

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/order.cpp:2906", "orderSelectedObjAdd", "Invalid player: %u", player: Exit Sub

    ' note that an order list graphic needs to be displayed
    bOrderEffectDisplayed = False

    it = apsDroidLists(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aDroid(psCurr).selected <> 0 Then
            If isBlueprint(psObj) Then
                objPos = SIMPLE_OBJECT_pos(psObj)
                If DROID_isConstructionDroid(psCurr) Then
                    ' Help build the planned structure.
                    orderDroidStatsLocDirAdd psCurr, DORDER_BUILD, aStruct(castStructure(psObj)).pStructureType, objPos.x, objPos.y, aStruct(castStructure(psObj)).rot.direction, add
                Else
                    ' Help watch the structure being built.
                    orderDroidLocAdd psCurr, DORDER_MOVE, objPos.x, objPos.y, add
                End If
                GoTo continue_
            End If

            order = chooseOrderObj(psCurr, psObj, specialOrderKeyDown())
            ' see if the order can be added to the list
            If order.typ <> DORDER_NONE Then
                If Not orderDroidObjAdd(psCurr, order, add) Then
                    ' if not just do it straight off
                    orderDroidObj psCurr, order.typ, order.psObj, ModeQueue
                End If
            End If
        End If
continue_:
        it = lnNext(it)
    Loop

    orderPlayOrderObjAudio player, psObj
End Sub

' This function just calls orderSelectedObjAdd with add = false.
Public Sub orderSelectedObj(ByVal player As Long, ByVal psObj As Long)
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/order.cpp:2956", "orderSelectedObj", "Invalid player: %u", player: Exit Sub
    orderSelectedObjAdd player, psObj, False
End Sub

' Given a player, this function send an order with localization and status to selected droids.
' If add is true, the orders are queued.
' @todo this function runs through all the player's droids, but only uses the selected ones and the ones that are construction droids. Consider an efficiency improvement.
Public Sub orderSelectedStatsLocDir(ByVal player As Long, ByVal order As DroidOrderType, ByVal psStats As Long, ByVal x As Long, ByVal y As Long, ByVal direction As Long, ByVal add As Boolean)
    Dim it As Long, psCurr As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/order.cpp:2967", "orderSelectedStatsLocDir", "Invalid player: %u", player: Exit Sub

    it = apsDroidLists(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aDroid(psCurr).selected <> 0 And DROID_isConstructionDroid(psCurr) Then
            If add Then
                orderDroidStatsLocDirAdd psCurr, order, psStats, x, y, direction
            Else
                orderDroidStatsLocDir psCurr, order, psStats, x, y, direction, ModeQueue
            End If
        End If
        it = lnNext(it)
    Loop
End Sub

' Same as orderSelectedStatsLocDir() but with two locations.
' @todo this function runs through all the player's droids, but only uses the selected ones. Consider an efficiency improvement.
Public Sub orderSelectedStatsTwoLocDir(ByVal player As Long, ByVal order As DroidOrderType, ByVal psStats As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal direction As Long, ByVal add As Boolean)
    Dim it As Long, psCurr As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/order.cpp:2991", "orderSelectedStatsTwoLocDir", "Invalid player: %u", player: Exit Sub

    it = apsDroidLists(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aDroid(psCurr).selected <> 0 Then
            If add Then
                orderDroidStatsTwoLocDirAdd psCurr, order, psStats, x1, y1, x2, y2, direction
            Else
                orderDroidStatsTwoLocDir psCurr, order, psStats, x1, y1, x2, y2, direction, ModeQueue
            End If
        End If
        it = lnNext(it)
    Loop
End Sub

' This function runs though all player's droids to check if any of then is a transporter. Returns the transporter droid if any was found, and NULL else.
Public Function FindATransporter(ByVal embarkee As Long) As Long
    Dim isCyborg As Boolean, bestDroid As Long, bestDist As Double, it As Long, psDroid As Long, dist As Double

    isCyborg = DROID_isCyborg(embarkee)

    bestDroid = 0
    bestDist = 4294967295#   ' ~0u

    it = apsDroidLists(aDroid(embarkee).player).head
    Do While it <> 0
        psDroid = lnObj(it)
        If (isCyborg And aDroid(psDroid).droidType = DROID_TRANSPORTER) Or aDroid(psDroid).droidType = DROID_SUPERTRANSPORTER Then
            dist = iHypot(aDroid(psDroid).pos.x - aDroid(embarkee).pos.x, aDroid(psDroid).pos.y - aDroid(embarkee).pos.y)
            If Not checkTransporterSpace(psDroid, embarkee, False) Then
                dist = U32(FromU32(dist + &H8000000))  ' Should prefer transports that aren't full.
            End If
            If dist < bestDist Then
                bestDroid = psDroid
                bestDist = dist
            End If
        End If
        it = lnNext(it)
    Loop

    FindATransporter = bestDroid
End Function

' Given a factory type, this function runs though all player's structures to check if any is of factory type. Returns the structure if any was found, and NULL else.
Private Function FindAFactory(ByVal player As Long, ByVal factoryType As Long) As Long
    Dim it As Long, psStruct As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/order.cpp:3042", "FindAFactory", "Invalid player: %u", player: FindAFactory = 0: Exit Function

    it = apsStructLists(player).head
    Do While it <> 0
        psStruct = lnObj(it)
        If asStructureStats(aStruct(psStruct).pStructureType).typ = factoryType Then
            FindAFactory = psStruct
            Exit Function
        End If
        it = lnNext(it)
    Loop

    FindAFactory = 0
End Function

' This function runs though all player's structures to check if any of then is a repair facility. Returns the structure if any was found, and NULL else.
Private Function FindARepairFacility(ByVal player As Long) As Long
    Dim it As Long, psStruct As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/order.cpp:3059", "FindARepairFacility", "Invalid player: %u", player: FindARepairFacility = 0: Exit Function

    it = apsStructLists(player).head
    Do While it <> 0
        psStruct = lnObj(it)
        If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REPAIR_FACILITY Then
            FindARepairFacility = psStruct
            Exit Function
        End If
        it = lnNext(it)
    Loop

    FindARepairFacility = 0
End Function

' This function returns true if the droid supports the secondary order, and false if not.
Public Function secondarySupported(ByVal psDroid As Long, ByVal sec As SECONDARY_ORDER) As Boolean
    Dim supported As Boolean, i As Long, weaponStats As Long, fallThrough As Boolean

    supported = True ' Default to supported.

    Select Case sec
    Case DSO_ASSIGN_PRODUCTION, DSO_ASSIGN_CYBORG_PRODUCTION, DSO_ASSIGN_VTOL_PRODUCTION, DSO_CLEAR_PRODUCTION, DSO_FIRE_DESIGNATOR
        ' DSO_CLEAR_PRODUCTION: remove production from a command droid
        If aDroid(psDroid).droidType <> DROID_COMMAND Then
            supported = False
        End If
        If (sec = DSO_ASSIGN_PRODUCTION And FindAFactory(aDroid(psDroid).player, REF_FACTORY) = 0) Or _
           (sec = DSO_ASSIGN_CYBORG_PRODUCTION And FindAFactory(aDroid(psDroid).player, REF_CYBORG_FACTORY) = 0) Or _
           (sec = DSO_ASSIGN_VTOL_PRODUCTION And FindAFactory(aDroid(psDroid).player, REF_VTOL_FACTORY) = 0) Then
            supported = False
        End If
        ' don't allow factories to be assigned to commanders during a Limbo Expand mission
        If (sec = DSO_ASSIGN_PRODUCTION Or sec = DSO_ASSIGN_CYBORG_PRODUCTION Or sec = DSO_ASSIGN_VTOL_PRODUCTION) _
           And missionLimboExpand() Then
            supported = False
        End If

    Case DSO_ATTACK_RANGE, DSO_ATTACK_LEVEL
        If sec = DSO_ATTACK_RANGE Then
            If aDroid(psDroid).droidType = DROID_SENSOR Then
                supported = False
            End If
            ' don't show the range levels if the droid doesn't have a weapon with different ranges
            If aDroid(psDroid).numWeaps > 0 Then
                For i = 0 To aDroid(psDroid).numWeaps - 1
                    weaponStats = DROID_getWeaponStats(psDroid, i)

                    If proj_GetLongRange(asWeaponStats(weaponStats), aDroid(psDroid).player) = proj_GetShortRange(asWeaponStats(weaponStats), aDroid(psDroid).player) Then
                        supported = False
                    Else
                        supported = True
                        Exit For
                    End If
                Next
            End If
            ' fall-through
        End If

        If aDroid(psDroid).droidType = DROID_REPAIR Or aDroid(psDroid).droidType = DROID_CYBORG_REPAIR Then
            supported = False
        End If
        If aDroid(psDroid).droidType = DROID_CONSTRUCT Or aDroid(psDroid).droidType = DROID_CYBORG_CONSTRUCT Then
            supported = False
        End If
        If aDroid(psDroid).droidType = DROID_ECM Then
            supported = False
        ElseIf objRadarDetector(psDroid) Then
            supported = False
        End If

    Case DSO_CIRCLE
        If Not DROID_isVtol(psDroid) Then
            supported = False
        End If

    Case DSO_ACCEPT_RETREP
        If aDroid(psDroid).droidType <> DROID_REPAIR And aDroid(psDroid).droidType <> DROID_CYBORG_REPAIR Then
            supported = False
        End If

    Case DSO_REPAIR_LEVEL, DSO_PATROL, DSO_HALTTYPE, DSO_RETURN_TO_LOC

    Case DSO_RECYCLE ' Only if player has got a factory.
        If (FindAFactory(aDroid(psDroid).player, REF_FACTORY) = 0) And _
           (FindAFactory(aDroid(psDroid).player, REF_CYBORG_FACTORY) = 0) And _
           (FindAFactory(aDroid(psDroid).player, REF_VTOL_FACTORY) = 0) And _
           (FindARepairFacility(aDroid(psDroid).player) = 0) Then
            supported = False
        End If

    Case Else
        supported = False
    End Select

    secondarySupported = supported
End Function

' This function returns the droid order's secondary state of the secondary order.
Public Function secondaryGetState(ByVal psDroid As Long, ByVal sec As SECONDARY_ORDER, Optional ByVal mode As QUEUE_MODE = ModeImmediate) As SECONDARY_STATE
    Dim state As Long

    state = aDroid(psDroid).secondaryOrder

    If mode = ModeQueue Then
        state = aDroid(psDroid).secondaryOrderPending  ' The UI wants to know the state, so return what the state will be after orders are synchronised.
    End If

    Select Case sec
    Case DSO_ATTACK_RANGE
        secondaryGetState = (state And DSS_ARANGE_MASK)
        Exit Function
    Case DSO_REPAIR_LEVEL
        secondaryGetState = (state And DSS_REPLEV_MASK)
        Exit Function
    Case DSO_ATTACK_LEVEL
        secondaryGetState = (state And DSS_ALEV_MASK)
        Exit Function
    Case DSO_ASSIGN_PRODUCTION, DSO_ASSIGN_CYBORG_PRODUCTION, DSO_ASSIGN_VTOL_PRODUCTION
        secondaryGetState = (state And DSS_ASSPROD_MASK)
        Exit Function
    Case DSO_RECYCLE
        secondaryGetState = (state And DSS_RECYCLE_MASK)
        Exit Function
    Case DSO_PATROL
        secondaryGetState = (state And DSS_PATROL_MASK)
        Exit Function
    Case DSO_CIRCLE
        secondaryGetState = (state And DSS_CIRCLE_MASK)
        Exit Function
    Case DSO_ACCEPT_RETREP
        secondaryGetState = (state And DSS_ACCREP_MASK)
        Exit Function
    Case DSO_HALTTYPE
        If aDroid(psDroid).order.typ = DORDER_HOLD Then
            secondaryGetState = DSS_HALT_HOLD
            Exit Function
        End If
        secondaryGetState = (state And DSS_HALT_MASK)
        Exit Function
    Case DSO_RETURN_TO_LOC
        secondaryGetState = (state And DSS_RTL_MASK)
        Exit Function
    Case DSO_FIRE_DESIGNATOR
        If cmdDroidGetDesignator(aDroid(psDroid).player) = psDroid Then
            secondaryGetState = DSS_FIREDES_SET
            Exit Function
        End If
    Case Else
    End Select

    secondaryGetState = DSS_NONE
End Function

' This function returns true if the droid needs repair according to the repair state, and in case there are some other droids selected, deselect those that are going to repair.
' @todo there is some problem related with the values of REPAIRLEV_HIGH and REPAIRLEV_LOW that needs to be fixed.
Private Function secondaryCheckDamageLevelDeselect(ByVal psDroid As Long, ByVal repairState_ As SECONDARY_STATE) As Boolean
    Dim repairLevel As Long, it As Long, psTempDroid As Long

    Select Case repairState_
    Case DSS_REPLEV_LOW: repairLevel = REPAIRLEV_HIGH  ' LOW -> HIGH, seems DSS_REPLEV_LOW and DSS_REPLEV_HIGH are badly named?
    Case DSS_REPLEV_HIGH: repairLevel = REPAIRLEV_LOW
    Case Else: repairLevel = 0                         ' DSS_REPLEV_NEVER
    End Select

    ' psDroid->body / psDroid->originalBody < repairLevel / 100, without integer truncation
    If Not ULT(MUL32(repairLevel, aDroid(psDroid).originalBody), MUL32(aDroid(psDroid).body, 100)) Then
        ' Only deselect the droid if there is another droid selected.
        If aDroid(psDroid).selected <> 0 And ULT(selectedPlayer, MAX_PLAYERS) Then
            it = apsDroidLists(selectedPlayer).head
            Do While it <> 0
                psTempDroid = lnObj(it)
                If psTempDroid <> psDroid And aDroid(psTempDroid).selected <> 0 Then
                    DeSelectDroid psDroid
                    Exit Do
                End If
                it = lnNext(it)
            Loop
        End If
        secondaryCheckDamageLevelDeselect = True
        Exit Function
    End If
    secondaryCheckDamageLevelDeselect = False
End Function

' This function checks the droid damage level against its secondary state. If the damage level is too high, then it sends an order to the droid to return to repair.
Public Sub secondaryCheckDamageLevel(ByVal psDroid As Long)
    Dim result As RtrBestResult, c As Boolean

    If secondaryCheckDamageLevelDeselect(psDroid, secondaryGetState(psDroid, DSO_REPAIR_LEVEL)) Then
        If Not DROID_isVtol(psDroid) Then
            If aDroid(psDroid).group <> UBYTE_MAX Then
                If Not (aDroid(psDroid).repairGroup = UBYTE_MAX) Then ASSERT_FAILURE "psDroid->repairGroup == 255", "src/order.cpp:3325", "secondaryCheckDamageLevel", "repairGroup was already: %u; setting to: %u", aDroid(psDroid).repairGroup, aDroid(psDroid).group
                aDroid(psDroid).repairGroup = aDroid(psDroid).group
                intGroupsChanged aDroid(psDroid).group
            End If
            aDroid(psDroid).group = UBYTE_MAX
        End If

        ' set return to repair if not on hold
        ' (&&: isVtolRearming, whose CHECK_DROID may log, only when the order is neither)
        c = aDroid(psDroid).order.typ <> DORDER_RTR And aDroid(psDroid).order.typ <> DORDER_RTB
        If c Then c = Not DROID_isVtolRearming(psDroid)
        If c Then
            If DROID_isVtol(psDroid) Then
                moveToRearm psDroid
            Else
                result = decideWhereToRepairAndBalance(psDroid)
                If result.typ = RTR_TYPE_REPAIR_FACILITY Then
                    If Not (result.psObj <> 0) Then ASSERT_FAILURE "result.psObj != nullptr", "src/order.cpp:3346", "secondaryCheckDamageLevel", "RTR_FACILITY but target is null"
                    orderDroidObj psDroid, DORDER_RTR, result.psObj, ModeQueue
                    Exit Sub
                ElseIf result.typ = RTR_TYPE_HQ Then
                    If Not (result.psObj <> 0) Then ASSERT_FAILURE "result.psObj != nullptr", "src/order.cpp:3352", "secondaryCheckDamageLevel", "RTR_TYPE_HQ but target is null"
                    orderDroid psDroid, DORDER_RTB, ModeQueue
                    Exit Sub
                ElseIf result.typ = RTR_TYPE_NO_RESULT Then
                    orderDroid psDroid, DORDER_RTB, ModeQueue
                    Exit Sub
                ElseIf result.typ = RTR_TYPE_DROID Then
                    If Not (result.psObj <> 0) Then ASSERT_FAILURE "result.psObj != nullptr", "src/order.cpp:3363", "secondaryCheckDamageLevel", "RTR_DROID but target is null"
                    orderDroidObj psDroid, DORDER_RTR, result.psObj, ModeQueue
                End If

            End If
        End If
    End If
End Sub

' (VB) decideWhereToRepairAndBalance's vectors' push_back
Private Sub pushLong(v() As Long, n As Long, ByVal x As Long)
    If n = 0 Then
        ReDim v(0 To 15)
    ElseIf n > UBound(v) Then
        ReDim Preserve v(0 To 2 * n - 1)
    End If
    v(n) = x
    n = n + 1
End Sub

Private Sub pushPos(v() As Vector3i, n As Long, x As Vector3i)
    If n = 0 Then
        ReDim v(0 To 15)
    ElseIf n > UBound(v) Then
        ReDim Preserve v(0 To 2 * n - 1)
    End If
    v(n) = x
    n = n + 1
End Sub

' balance the load at random
' always prefer faster repairs
Private Function decideWhereToRepairAndBalance(ByVal psDroid As Long) As RtrBestResult
    Dim bestDistToRepairFac As Long, bestDistToRepairDroid As Long, thisDistToRepair As Long, psHq As Long
    Dim bestDroidPos As Vector3i, bestFacPos As Vector3i, it As Long, psStruct As Long, psCurr As Long
    Dim bestRepairPoint As Vector3i, i As Long, diff As Vector2i, which As Long, useMission As Boolean

    bestDistToRepairFac = INT32_MAX
    bestDistToRepairDroid = INT32_MAX
    thisDistToRepair = 0
    psHq = 0
    ' clear vectors from previous invocations
    vFacilityPos_size = 0
    vFacility_size = 0
    vFacilityCloseEnough_size = 0
    vDroidCloseEnough_size = 0
    vDroidPos_size = 0
    vDroid_size = 0

    it = apsStructLists(aDroid(psDroid).player).head
    Do While it <> 0
        psStruct = lnObj(it)
        If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_HQ Then
            If CheckInScrollLimits(aStruct(psStruct).pos.x, aStruct(psStruct).pos.y) Then
                psHq = psStruct
            End If
            GoTo continueStruct
        End If
        If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REPAIR_FACILITY And aStruct(psStruct).status = SS_BUILT Then
            If Not CheckInScrollLimits(aStruct(psStruct).pos.x, aStruct(psStruct).pos.y) Then
                GoTo continueStruct
            End If
            thisDistToRepair = droidSqDist(psDroid, psStruct)
            If thisDistToRepair <= 0 Then
                GoTo continueStruct  ' cannot reach position
            End If
            pushPos vFacilityPos, vFacilityPos_size, aStruct(psStruct).pos
            pushLong vFacility, vFacility_size, psStruct
            If bestDistToRepairFac > thisDistToRepair Then
                bestDistToRepairFac = thisDistToRepair
                bestFacPos = aStruct(psStruct).pos
            End If
        End If
continueStruct:
        it = lnNext(it)
    Loop

    ' If we are repair droid ourselves that accept retreating units, don't consider
    ' other repairs droids. Since this would cause traffic jams if we are attacked amongst
    ' other accepting repair droids. Thus causing chaos as repair units simply clump up
    ' instead of actually retreating.
    If Not ((aDroid(psDroid).droidType = DROID_REPAIR Or aDroid(psDroid).droidType = DROID_CYBORG_REPAIR) _
            And secondaryGetState(psDroid, DSO_ACCEPT_RETREP) <> 0) Then
        ' one of these lists is empty when on mission
        useMission = apsDroidLists(aDroid(psDroid).player).size = 0
        If useMission Then
            it = mission_apsDroidLists(aDroid(psDroid).player).head
        Else
            it = apsDroidLists(aDroid(psDroid).player).head
        End If
        Do While it <> 0
            psCurr = lnObj(it)
            ' Accept any repair droids that accept retreating units
            If (aDroid(psCurr).droidType = DROID_REPAIR Or aDroid(psCurr).droidType = DROID_CYBORG_REPAIR) _
               And secondaryGetState(psCurr, DSO_ACCEPT_RETREP) <> 0 Then
                If Not CheckInScrollLimits(aDroid(psCurr).pos.x, aDroid(psCurr).pos.y) Then
                    GoTo continueDroid
                End If
                thisDistToRepair = droidSqDist(psDroid, psCurr)
                If thisDistToRepair <= 0 Then
                    GoTo continueDroid ' unreachable
                End If
                pushPos vDroidPos, vDroidPos_size, aDroid(psCurr).pos
                pushLong vDroid, vDroid_size, psCurr
                If bestDistToRepairDroid > thisDistToRepair Then
                    bestDistToRepairDroid = thisDistToRepair
                    bestDroidPos = aDroid(psCurr).pos
                End If
            End If
continueDroid:
            it = lnNext(it)
        Loop
    End If

    If Not (bestDistToRepairFac > 0) Then ASSERT_FAILURE "bestDistToRepairFac > 0", "src/order.cpp:3463", "decideWhereToRepairAndBalance", "Bad distance to repair facility"
    If Not (bestDistToRepairDroid > 0) Then ASSERT_FAILURE "bestDistToRepairDroid > 0", "src/order.cpp:3464", "decideWhereToRepairAndBalance", "Bad distance to repair droid"
    ' debug(LOG_INFO, "found a total of %lu RT, and %lu RF", vDroid.size(), vFacility.size());

    ' the center of this area starts at the closest repair droid/facility!
    If bestDistToRepairFac < bestDistToRepairDroid Then bestRepairPoint = bestFacPos Else bestRepairPoint = bestDroidPos
    ' find all close enough repairing candidates
    For i = 0 To vFacilityPos_size - 1
        diff.x = bestRepairPoint.x - vFacilityPos(i).x
        diff.y = bestRepairPoint.y - vFacilityPos(i).y
        If dot2i(diff, diff) <= MAGIC_SUITABLE_REPAIR_AREA Then
            pushLong vFacilityCloseEnough, vFacilityCloseEnough_size, i
        End If
    Next
    For i = 0 To vDroidPos_size - 1
        diff.x = bestRepairPoint.x - vDroidPos(i).x
        diff.y = bestRepairPoint.y - vDroidPos(i).y
        If dot2i(diff, diff) <= MAGIC_SUITABLE_REPAIR_AREA Then
            pushLong vDroidCloseEnough, vDroidCloseEnough_size, i
        End If
    Next

    ' debug(LOG_INFO, "found  %lu RT, and %lu RF in suitable area", vDroidCloseEnough.size(), vFacilityCloseEnough.size());
    ' prefer facilities, they re much more efficient than droids
    If vFacilityCloseEnough_size = 1 Then
        decideWhereToRepairAndBalance = RtrBestResult_(RTR_TYPE_REPAIR_FACILITY, vFacility(vFacilityCloseEnough(0)))
        Exit Function

    ElseIf vFacilityCloseEnough_size > 1 Then
        which = gameRand(vFacilityCloseEnough_size)
        decideWhereToRepairAndBalance = RtrBestResult_(RTR_TYPE_REPAIR_FACILITY, vFacility(vFacilityCloseEnough(which)))
        Exit Function
    End If

    ' no facilities :( fallback on droids
    If vDroidCloseEnough_size = 1 Then
        decideWhereToRepairAndBalance = RtrBestResult_(RTR_TYPE_DROID, vDroid(vDroidCloseEnough(0)))
        Exit Function
    ElseIf vDroidCloseEnough_size > 1 Then
        which = gameRand(vDroidCloseEnough_size)
        decideWhereToRepairAndBalance = RtrBestResult_(RTR_TYPE_DROID, vDroid(vDroidCloseEnough(which)))
        Exit Function
    End If

    ' go to headquarters, if any
    If psHq <> 0 Then
        decideWhereToRepairAndBalance = RtrBestResult_(RTR_TYPE_HQ, psHq)
        Exit Function
    End If
    ' screw it
    decideWhereToRepairAndBalance = RtrBestResult_(RTR_TYPE_NO_RESULT, 0)
End Function

' This function assigns a state to a droid. It returns true if it assigned and false if it failed to assign.
' Note that this also modifies primary order in some cases
Public Function secondarySetState(ByVal psDroid As Long, ByVal sec As SECONDARY_ORDER, ByVal State As SECONDARY_STATE, Optional ByVal mode As QUEUE_MODE = ModeQueue) As Boolean
    Dim CurrState As Long, factType As Long, prodType As Long, factoryInc As Long, retVal As Boolean
    Dim psTransport As Long, order As DroidOrderType, secondaryMask As Long, secondarySet As Long
    Dim newSecondaryState As Long, it As Long, psStruct As Long, nxt As Long, psCurr As Long, prevGroup As Long

    CurrState = aDroid(psDroid).secondaryOrder
    If bMultiMessages And mode = ModeQueue Then
        CurrState = aDroid(psDroid).secondaryOrderPending
    End If

    ' Figure out what the new secondary state will be (once the order is synchronised.
    ' Why does this have to be so ridiculously complicated?
    secondaryMask = 0
    secondarySet = 0

    Select Case sec
    Case DSO_ATTACK_RANGE
        secondaryMask = DSS_ARANGE_MASK
        secondarySet = State
    Case DSO_REPAIR_LEVEL
        secondaryMask = DSS_REPLEV_MASK
        secondarySet = State
    Case DSO_ATTACK_LEVEL
        secondaryMask = DSS_ALEV_MASK
        secondarySet = State
    Case DSO_ASSIGN_PRODUCTION
        If aDroid(psDroid).droidType = DROID_COMMAND Then
            secondaryMask = DSS_ASSPROD_FACT_MASK
            secondarySet = State And DSS_ASSPROD_MASK
        End If
    Case DSO_ASSIGN_CYBORG_PRODUCTION
        If aDroid(psDroid).droidType = DROID_COMMAND Then
            secondaryMask = DSS_ASSPROD_CYB_MASK
            secondarySet = State And DSS_ASSPROD_MASK
        End If
    Case DSO_ASSIGN_VTOL_PRODUCTION
        If aDroid(psDroid).droidType = DROID_COMMAND Then
            secondaryMask = DSS_ASSPROD_VTOL_MASK
            secondarySet = State And DSS_ASSPROD_MASK
        End If
    Case DSO_CLEAR_PRODUCTION
        If aDroid(psDroid).droidType = DROID_COMMAND Then
            secondaryMask = State And DSS_ASSPROD_MASK
        End If
    Case DSO_RECYCLE
        If (State And DSS_RECYCLE_MASK) <> 0 Then
            secondaryMask = DSS_RTL_MASK Or DSS_RECYCLE_MASK Or DSS_HALT_MASK
            secondarySet = DSS_RECYCLE_SET Or DSS_HALT_GUARD
        Else
            secondaryMask = DSS_RECYCLE_MASK
        End If
    Case DSO_CIRCLE  ' This doesn't even make any sense whatsoever as a secondary order...
        secondaryMask = DSS_CIRCLE_MASK
        If (State And DSS_CIRCLE_SET) <> 0 Then secondarySet = DSS_CIRCLE_SET Else secondarySet = 0
    Case DSO_PATROL  ' This doesn't even make any sense whatsoever as a secondary order...
        secondaryMask = DSS_PATROL_MASK
        If (State And DSS_PATROL_SET) <> 0 Then secondarySet = DSS_PATROL_SET Else secondarySet = 0
    Case DSO_HALTTYPE
        Select Case State And DSS_HALT_MASK
        Case DSS_HALT_PURSUE, DSS_HALT_GUARD, DSS_HALT_HOLD
            secondaryMask = DSS_HALT_MASK
            secondarySet = State
        End Select
    Case DSO_RETURN_TO_LOC
        secondaryMask = DSS_RTL_MASK
        Select Case State And DSS_RTL_MASK
        Case DSS_RTL_REPAIR, DSS_RTL_BASE
            secondarySet = State
        Case DSS_RTL_TRANSPORT
            psTransport = FindATransporter(psDroid)
            If psTransport <> 0 Then
                secondarySet = State
            End If
        End Select
        If (CurrState And DSS_HALT_MASK) = DSS_HALT_HOLD Then
            secondaryMask = secondaryMask Or DSS_HALT_MASK
            secondarySet = secondarySet Or DSS_HALT_GUARD
        End If
    Case DSO_ACCEPT_RETREP
        secondaryMask = DSS_ACCREP_MASK
        secondarySet = State
    Case DSO_UNUSED, DSO_FIRE_DESIGNATOR
        ' Do nothing.
    End Select
    newSecondaryState = (CurrState And Not secondaryMask) Or secondarySet

    If bMultiMessages And mode = ModeQueue Then
        If sec = DSO_REPAIR_LEVEL Then
            secondaryCheckDamageLevelDeselect psDroid, State  ' Deselect droid immediately, if applicable, so it isn't ordered around by mistake.
        End If

        sendDroidSecondary psDroid, sec, State
        aDroid(psDroid).secondaryOrderPending = newSecondaryState
        aDroid(psDroid).secondaryOrderPendingCount = aDroid(psDroid).secondaryOrderPendingCount + 1
        secondarySetState = True  ' Wait for our order before changing the droid.
        Exit Function
    End If


    ' set the state for any droids in the command group
    If (sec <> DSO_RECYCLE) And _
       aDroid(psDroid).droidType = DROID_COMMAND And _
       aDroid(psDroid).psGroup <> 0 Then
        If aGroup(aDroid(psDroid).psGroup).typ = GT_COMMAND Then
            DROID_GROUP_setSecondary aDroid(psDroid).psGroup, sec, State
        End If
    End If

    retVal = True
    Select Case sec
    Case DSO_ATTACK_RANGE
        CurrState = (CurrState And Not DSS_ARANGE_MASK) Or State

    Case DSO_REPAIR_LEVEL
        CurrState = (CurrState And Not DSS_REPLEV_MASK) Or State
        aDroid(psDroid).secondaryOrder = CurrState
        secondaryCheckDamageLevel psDroid

    Case DSO_ATTACK_LEVEL
        CurrState = (CurrState And Not DSS_ALEV_MASK) Or State
        If State = DSS_ALEV_NEVER Then
            If orderState(psDroid, DORDER_ATTACK) Then
                ' just kill these orders
                orderDroid psDroid, DORDER_STOP, ModeImmediate
                If DROID_isVtol(psDroid) Then
                    moveToRearm psDroid
                End If
            ElseIf DROID_isAttacking(psDroid) Then
                ' send the unit back to the guard position
                actionDroid psDroid, DACTION_NONE
            ElseIf orderState(psDroid, DORDER_PATROL) Then
                ' send the unit back to the patrol
                actionDroid_loc psDroid, DACTION_RETURNTOPOS, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y
            End If
        End If


    Case DSO_ASSIGN_PRODUCTION, DSO_ASSIGN_CYBORG_PRODUCTION, DSO_ASSIGN_VTOL_PRODUCTION
        If sec = DSO_ASSIGN_PRODUCTION Then
            prodType = REF_FACTORY
        ElseIf sec = DSO_ASSIGN_CYBORG_PRODUCTION Then
            prodType = REF_CYBORG_FACTORY
        Else
            prodType = REF_VTOL_FACTORY
        End If

        If aDroid(psDroid).droidType = DROID_COMMAND Then
            ' look for the factories
            it = apsStructLists(aDroid(psDroid).player).head
            Do While it <> 0
                psStruct = lnObj(it)
                factType = asStructureStats(aStruct(psStruct).pStructureType).typ
                If factType = REF_FACTORY Or _
                   factType = REF_VTOL_FACTORY Or _
                   factType = REF_CYBORG_FACTORY Then
                    factoryInc = aFlagPos(aStruct(psStruct).functionality.factory.psAssemblyPoint).factoryInc
                    If factType = REF_FACTORY Then
                        factoryInc = factoryInc + DSS_ASSPROD_SHIFT
                    ElseIf factType = REF_CYBORG_FACTORY Then
                        factoryInc = factoryInc + DSS_ASSPROD_CYBORG_SHIFT
                    Else
                        factoryInc = factoryInc + DSS_ASSPROD_VTOL_SHIFT
                    End If
                    If (CurrState And SHL(1, factoryInc)) = 0 And _
                       (State And SHL(1, factoryInc)) <> 0 Then
                        assignFactoryCommandDroid psStruct, psDroid ' assign this factory to the command droid
                    ElseIf (prodType = factType) And _
                           (CurrState And SHL(1, factoryInc)) <> 0 And _
                           (State And SHL(1, factoryInc)) = 0 Then
                        ' remove this factory from the command droid
                        assignFactoryCommandDroid psStruct, 0
                    End If
                End If
                it = lnNext(it)
            Loop
            If prodType = REF_FACTORY Then
                CurrState = CurrState And Not DSS_ASSPROD_FACT_MASK
            ElseIf prodType = REF_CYBORG_FACTORY Then
                CurrState = CurrState And Not DSS_ASSPROD_CYB_MASK
            Else
                CurrState = CurrState And Not DSS_ASSPROD_VTOL_MASK
            End If
            CurrState = CurrState Or (State And DSS_ASSPROD_MASK)
        End If

    Case DSO_CLEAR_PRODUCTION
        If aDroid(psDroid).droidType = DROID_COMMAND Then
            ' simply clear the flag - all the factory stuff is done in assignFactoryCommandDroid
            CurrState = CurrState And Not (State And DSS_ASSPROD_MASK)
        End If


    Case DSO_RECYCLE
        If (State And DSS_RECYCLE_MASK) <> 0 Then
            If Not orderState(psDroid, DORDER_RECYCLE) Then
                orderDroid psDroid, DORDER_RECYCLE, ModeImmediate
            End If
            CurrState = CurrState And Not (DSS_RTL_MASK Or DSS_RECYCLE_MASK Or DSS_HALT_MASK)
            CurrState = CurrState Or DSS_RECYCLE_SET Or DSS_HALT_GUARD
            prevGroup = aDroid(psDroid).group
            aDroid(psDroid).group = UBYTE_MAX
            aDroid(psDroid).repairGroup = UBYTE_MAX
            intGroupsChanged prevGroup
            If aDroid(psDroid).psGroup <> 0 Then
                If aDroid(psDroid).droidType = DROID_COMMAND Then
                    ' remove all the units from the commanders group
                    it = aGroup(aDroid(psDroid).psGroup).psList.head
                    Do While it <> 0
                        nxt = lnNext(it)
                        psCurr = lnObj(it)
                        DROID_GROUP_remove aDroid(psCurr).psGroup, psCurr
                        orderDroid psCurr, DORDER_STOP, ModeImmediate
                        it = nxt
                    Loop
                ElseIf aGroup(aDroid(psDroid).psGroup).typ = GT_COMMAND Then
                    DROID_GROUP_remove aDroid(psDroid).psGroup, psDroid
                End If
            End If
        Else
            If orderState(psDroid, DORDER_RECYCLE) Then
                orderDroid psDroid, DORDER_STOP, ModeImmediate
            End If
            CurrState = CurrState And Not DSS_RECYCLE_MASK
        End If
    Case DSO_CIRCLE
        If (State And DSS_CIRCLE_SET) <> 0 Then
            CurrState = CurrState Or DSS_CIRCLE_SET
        Else
            CurrState = CurrState And Not DSS_CIRCLE_MASK
        End If
    Case DSO_ACCEPT_RETREP
        CurrState = (CurrState And Not DSS_ACCREP_MASK) Or State
    Case DSO_PATROL
        If (State And DSS_PATROL_SET) <> 0 Then
            CurrState = CurrState Or DSS_PATROL_SET
        Else
            CurrState = CurrState And Not DSS_PATROL_MASK
        End If
    Case DSO_HALTTYPE
        Select Case State And DSS_HALT_MASK
        Case DSS_HALT_PURSUE
            CurrState = CurrState And Not DSS_HALT_MASK
            CurrState = CurrState Or DSS_HALT_PURSUE
            If orderState(psDroid, DORDER_GUARD) Then
                orderDroid psDroid, DORDER_STOP, ModeImmediate
            End If
        Case DSS_HALT_GUARD
            CurrState = CurrState And Not DSS_HALT_MASK
            CurrState = CurrState Or DSS_HALT_GUARD
            orderDroidLoc psDroid, DORDER_GUARD, aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, ModeImmediate
        Case DSS_HALT_HOLD
            CurrState = CurrState And Not DSS_HALT_MASK
            CurrState = CurrState Or DSS_HALT_HOLD
            If Not orderState(psDroid, DORDER_FIRESUPPORT) Then
                orderDroid psDroid, DORDER_STOP, ModeImmediate
            End If
        End Select
    Case DSO_RETURN_TO_LOC
        If (State And DSS_RTL_MASK) = 0 Then
            If orderState(psDroid, DORDER_RTR) Or _
               orderState(psDroid, DORDER_RTB) Or _
               orderState(psDroid, DORDER_EMBARK) Then
                orderDroid psDroid, DORDER_STOP, ModeImmediate
            End If
            CurrState = CurrState And Not DSS_RTL_MASK
        Else
            order = DORDER_NONE
            CurrState = CurrState And Not DSS_RTL_MASK
            If (CurrState And DSS_HALT_MASK) = DSS_HALT_HOLD Then
                CurrState = CurrState And Not DSS_HALT_MASK
                CurrState = CurrState Or DSS_HALT_GUARD
            End If
            Select Case State And DSS_RTL_MASK
            Case DSS_RTL_REPAIR
                order = DORDER_RTR
                CurrState = CurrState Or DSS_RTL_REPAIR
                ' can't clear the selection here cos it breaks the secondary order screen
            Case DSS_RTL_BASE
                order = DORDER_RTB
                CurrState = CurrState Or DSS_RTL_BASE
            Case DSS_RTL_TRANSPORT
                psTransport = FindATransporter(psDroid)
                If psTransport <> 0 Then
                    order = DORDER_EMBARK
                    CurrState = CurrState Or DSS_RTL_TRANSPORT
                    If Not orderState(psDroid, DORDER_EMBARK) Then
                        orderDroidObj psDroid, DORDER_EMBARK, psTransport, ModeImmediate
                    End If
                Else
                    retVal = False
                End If
            Case Else
                order = DORDER_NONE
            End Select
            If Not orderState(psDroid, order) Then
                orderDroid psDroid, order, ModeImmediate
            End If
        End If

    Case DSO_FIRE_DESIGNATOR
        ' don't actually set any secondary flags - the cmdDroid array is
        ' always used to determine which commander is the designator
        If (State And DSS_FIREDES_SET) <> 0 Then
            cmdDroidSetDesignator psDroid
        ElseIf cmdDroidGetDesignator(aDroid(psDroid).player) = psDroid Then
            cmdDroidClearDesignator aDroid(psDroid).player
        End If

    Case Else
    End Select

    If CurrState <> newSecondaryState Then
        debug_ 3957, LOG_WARNING, "secondarySetState", "Guessed the new secondary state incorrectly, expected 0x%08X, got 0x%08X, was 0x%08X, sec = %d, state = 0x%08X.", newSecondaryState, CurrState, aDroid(psDroid).secondaryOrder, sec, State
    End If
    aDroid(psDroid).secondaryOrder = CurrState
    aDroid(psDroid).secondaryOrderPendingCount = MAX_(aDroid(psDroid).secondaryOrderPendingCount - 1, 0)
    If aDroid(psDroid).secondaryOrderPendingCount = 0 Then
        aDroid(psDroid).secondaryOrderPending = aDroid(psDroid).secondaryOrder  ' If no orders are pending, make sure UI uses the actual state.
    End If

    secondarySetState = retVal
End Function

' This function assigns all droids of the group to the state.
' @todo this function runs through all the player's droids. Consider something more efficient to select a group.
' @todo SECONDARY_STATE argument is called "state", which is not current style. Suggestion to change it to "pState".
Private Sub secondarySetGroupState(ByVal player As Long, ByVal group As Long, ByVal sec As SECONDARY_ORDER, ByVal state As SECONDARY_STATE)
    Dim it As Long, psCurr As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/order.cpp:3975", "secondarySetGroupState", "Invalid player: %u", player: Exit Sub

    it = apsDroidLists(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aDroid(psCurr).group = group And _
           secondaryGetState(psCurr, sec) <> state Then
            secondarySetState psCurr, sec, state
        End If
        it = lnNext(it)
    Loop
End Sub

' This function returns the average secondary state of a numerical group of a player.
' @todo this function runs through all the player's droids. Consider something more efficient to select a group.
' @todo this function uses a "local" define. Consider removing it, refactoring this function.
Private Function secondaryGetAverageGroupState(ByVal player As Long, ByVal group As Long, ByVal mask As Long) As SECONDARY_STATE
    Const MAX_STATES As Long = 5
    Dim aStateCount_state(0 To MAX_STATES - 1) As Long, aStateCount_num(0 To MAX_STATES - 1) As Long
    Dim i As Long, numStates As Long, maxI As Long, it As Long, psCurr As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/order.cpp:3994", "secondaryGetAverageGroupState", "Invalid player: %u", player: secondaryGetAverageGroupState = DSS_NONE: Exit Function

    ' count the number of units for each state
    numStates = 0
    it = apsDroidLists(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aDroid(psCurr).group = group Then
            For i = 0 To numStates - 1
                If aStateCount_state(i) = (aDroid(psCurr).secondaryOrder And mask) Then
                    aStateCount_num(i) = aStateCount_num(i) + 1
                    Exit For
                End If
            Next

            If i = numStates Then
                aStateCount_state(numStates) = aDroid(psCurr).secondaryOrder And mask
                aStateCount_num(numStates) = 1
                numStates = numStates + 1
            End If
        End If
        it = lnNext(it)
    Loop

    maxI = 0
    For i = 0 To numStates - 1
        If aStateCount_num(i) > aStateCount_num(maxI) Then
            maxI = i
        End If
    Next

    secondaryGetAverageGroupState = aStateCount_state(maxI)
End Function

' This function sets all the group's members to have the same secondary state as the average secondary state of the group.
' @todo this function runs through all the player's droids. Consider something more efficient to select a group.
' @todo this function uses a "local" define. Consider removing it, refactoring this function.
Public Sub secondarySetAverageGroupState(ByVal player As Long, ByVal group As Long)
    Const MAX_ORDERS As Long = 4
    Dim aOrders_order(0 To MAX_ORDERS - 1) As SECONDARY_ORDER, aOrders_mask(0 To MAX_ORDERS - 1) As Long
    Dim i As Long, state As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/order.cpp:4047", "secondarySetAverageGroupState", "Invalid player: %u", player: Exit Sub

    ' lookup table for orders and masks
    aOrders_order(0) = DSO_ATTACK_RANGE: aOrders_mask(0) = DSS_ARANGE_MASK
    aOrders_order(1) = DSO_REPAIR_LEVEL: aOrders_mask(1) = DSS_REPLEV_MASK
    aOrders_order(2) = DSO_ATTACK_LEVEL: aOrders_mask(2) = DSS_ALEV_MASK
    aOrders_order(3) = DSO_HALTTYPE: aOrders_mask(3) = DSS_HALT_MASK

    For i = 0 To MAX_ORDERS - 1
        state = secondaryGetAverageGroupState(player, group, aOrders_mask(i))
        secondarySetGroupState player, group, aOrders_order(i), state
    Next
End Sub

' This function changes the structure's secondary state to be the function input's state.
' Returns true if the function changed the structure's state, and false if it did not.
' @todo SECONDARY_STATE argument is called "State", which is not current style. Suggestion to change it to "pState".
Public Function setFactoryState(ByVal psStruct As Long, ByVal sec As SECONDARY_ORDER, ByVal State As SECONDARY_STATE) As Boolean
    Dim CurrState As Long, retVal As Boolean, ok As Boolean

    ok = psStruct <> 0
    If ok Then ok = STRUCTURE_isFactory(psStruct)
    If Not ok Then
        ASSERT_FAILURE "false", "src/order.cpp:4080", "setFactoryState", "setFactoryState: structure is not a factory"
        setFactoryState = False
        Exit Function
    End If

    CurrState = aStruct(psStruct).functionality.factory.secondaryOrder

    retVal = True
    Select Case sec
    Case DSO_ATTACK_RANGE
        CurrState = (CurrState And Not DSS_ARANGE_MASK) Or State

    Case DSO_REPAIR_LEVEL
        CurrState = (CurrState And Not DSS_REPLEV_MASK) Or State

    Case DSO_ATTACK_LEVEL
        CurrState = (CurrState And Not DSS_ALEV_MASK) Or State

    Case DSO_PATROL
        If (State And DSS_PATROL_SET) <> 0 Then
            CurrState = CurrState Or DSS_PATROL_SET
        Else
            CurrState = CurrState And Not DSS_PATROL_MASK
        End If

    Case DSO_HALTTYPE
        Select Case State And DSS_HALT_MASK
        Case DSS_HALT_PURSUE
            CurrState = CurrState And Not DSS_HALT_MASK
            CurrState = CurrState Or DSS_HALT_PURSUE
        Case DSS_HALT_GUARD
            CurrState = CurrState And Not DSS_HALT_MASK
            CurrState = CurrState Or DSS_HALT_GUARD
        Case DSS_HALT_HOLD
            CurrState = CurrState And Not DSS_HALT_MASK
            CurrState = CurrState Or DSS_HALT_HOLD
        End Select
    Case Else
    End Select

    aStruct(psStruct).functionality.factory.secondaryOrder = CurrState

    setFactoryState = retVal
End Function

' This function sets the structure's secondary state to be pState.
'  return true except on an ASSERT (which is not a good design.)
'  or, an invalid factory.
Public Function getFactoryState(ByVal psStruct As Long, ByVal sec As SECONDARY_ORDER, pState As SECONDARY_STATE) As Boolean
    Dim state As Long, ok As Boolean

    ok = psStruct <> 0
    If ok Then ok = STRUCTURE_isFactory(psStruct)
    If Not ok Then ASSERT_FAILURE "psStruct && psStruct->isFactory()", "src/order.cpp:4147", "getFactoryState", "Structure is not a factory": getFactoryState = False: Exit Function
    If aStruct(psStruct).pFunctionality <> 0 Then
        state = aStruct(psStruct).functionality.factory.secondaryOrder

        Select Case sec
        Case DSO_ATTACK_RANGE
            pState = (state And DSS_ARANGE_MASK)
        Case DSO_REPAIR_LEVEL
            pState = (state And DSS_REPLEV_MASK)
        Case DSO_ATTACK_LEVEL
            pState = (state And DSS_ALEV_MASK)
        Case DSO_PATROL
            pState = (state And DSS_PATROL_MASK)
        Case DSO_HALTTYPE
            pState = (state And DSS_HALT_MASK)
        Case Else
            pState = 0
        End Select

        getFactoryState = True
        Exit Function
    End If
    getFactoryState = False
End Function

' lasSat structure can select a target
' @todo improve documentation: it is not clear what this function performs by the current documentation.
Public Sub orderStructureObj(ByVal player As Long, ByVal psObj As Long)
    Dim it As Long, psStruct As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/order.cpp:4185", "orderStructureObj", "Invalid player: %u", player: Exit Sub

    it = apsStructLists(player).head
    Do While it <> 0
        psStruct = lnObj(it)
        If lasSatStructSelected(psStruct) Then
            ' send the weapon fire
            sendLasSat player, psStruct, psObj

            Exit Do
        End If
        it = lnNext(it)
    Loop
End Sub

' This function maps the order enum to its name, returning its enum name as a (const char*)
' Formally, this function is equivalent to a stl map: for a given key (enum), returns a mapped value (char*).
Public Function getDroidOrderName(ByVal order As DroidOrderType) As String
    Select Case order
    Case DORDER_NONE: getDroidOrderName = "DORDER_NONE": Exit Function
    Case DORDER_STOP: getDroidOrderName = "DORDER_STOP": Exit Function
    Case DORDER_MOVE: getDroidOrderName = "DORDER_MOVE": Exit Function
    Case DORDER_ATTACK: getDroidOrderName = "DORDER_ATTACK": Exit Function
    Case DORDER_BUILD: getDroidOrderName = "DORDER_BUILD": Exit Function
    Case DORDER_HELPBUILD: getDroidOrderName = "DORDER_HELPBUILD": Exit Function
    Case DORDER_LINEBUILD: getDroidOrderName = "DORDER_LINEBUILD": Exit Function
    Case DORDER_DEMOLISH: getDroidOrderName = "DORDER_DEMOLISH": Exit Function
    Case DORDER_REPAIR: getDroidOrderName = "DORDER_REPAIR": Exit Function
    Case DORDER_OBSERVE: getDroidOrderName = "DORDER_OBSERVE": Exit Function
    Case DORDER_FIRESUPPORT: getDroidOrderName = "DORDER_FIRESUPPORT": Exit Function
    Case DORDER_UNUSED_4: getDroidOrderName = "DORDER_UNUSED_4": Exit Function
    Case DORDER_UNUSED_2: getDroidOrderName = "DORDER_UNUSED_2": Exit Function
    Case DORDER_RTB: getDroidOrderName = "DORDER_RTB": Exit Function
    Case DORDER_RTR: getDroidOrderName = "DORDER_RTR": Exit Function
    Case DORDER_UNUSED_5: getDroidOrderName = "DORDER_UNUSED_5": Exit Function
    Case DORDER_EMBARK: getDroidOrderName = "DORDER_EMBARK": Exit Function
    Case DORDER_DISEMBARK: getDroidOrderName = "DORDER_DISEMBARK": Exit Function
    Case DORDER_ATTACKTARGET: getDroidOrderName = "DORDER_ATTACKTARGET": Exit Function
    Case DORDER_COMMANDERSUPPORT: getDroidOrderName = "DORDER_COMMANDERSUPPORT": Exit Function
    Case DORDER_BUILDMODULE: getDroidOrderName = "DORDER_BUILDMODULE": Exit Function
    Case DORDER_RECYCLE: getDroidOrderName = "DORDER_RECYCLE": Exit Function
    Case DORDER_TRANSPORTOUT: getDroidOrderName = "DORDER_TRANSPORTOUT": Exit Function
    Case DORDER_TRANSPORTIN: getDroidOrderName = "DORDER_TRANSPORTIN": Exit Function
    Case DORDER_TRANSPORTRETURN: getDroidOrderName = "DORDER_TRANSPORTRETURN": Exit Function
    Case DORDER_GUARD: getDroidOrderName = "DORDER_GUARD": Exit Function
    Case DORDER_DROIDREPAIR: getDroidOrderName = "DORDER_DROIDREPAIR": Exit Function
    Case DORDER_RESTORE: getDroidOrderName = "DORDER_RESTORE": Exit Function
    Case DORDER_SCOUT: getDroidOrderName = "DORDER_SCOUT": Exit Function
    Case DORDER_UNUSED_3: getDroidOrderName = "DORDER_UNUSED_3": Exit Function
    Case DORDER_UNUSED: getDroidOrderName = "DORDER_UNUSED": Exit Function
    Case DORDER_PATROL: getDroidOrderName = "DORDER_PATROL": Exit Function
    Case DORDER_REARM: getDroidOrderName = "DORDER_REARM": Exit Function
    Case DORDER_RECOVER: getDroidOrderName = "DORDER_RECOVER": Exit Function
    Case DORDER_UNUSED_6: getDroidOrderName = "DORDER_UNUSED_6": Exit Function
    Case DORDER_RTR_SPECIFIED: getDroidOrderName = "DORDER_RTR_SPECIFIED": Exit Function
    Case DORDER_CIRCLE: getDroidOrderName = "DORDER_CIRCLE": Exit Function
    Case DORDER_HOLD: getDroidOrderName = "DORDER_HOLD": Exit Function
    End Select

    ASSERT_FAILURE "false", "src/order.cpp:4247", "getDroidOrderName", "DROID_ORDER out of range: %u", order

    getDroidOrderName = "DORDER_#INVALID#"
End Function

Public Function getDroidOrderKey(ByVal order As DroidOrderType) As String
    Select Case order
    Case DORDER_NONE: getDroidOrderKey = "N": Exit Function
    Case DORDER_STOP: getDroidOrderKey = "Stop": Exit Function
    Case DORDER_MOVE: getDroidOrderKey = "M": Exit Function
    Case DORDER_ATTACK: getDroidOrderKey = "A": Exit Function
    Case DORDER_BUILD: getDroidOrderKey = "B": Exit Function
    Case DORDER_HELPBUILD: getDroidOrderKey = "hB": Exit Function
    Case DORDER_LINEBUILD: getDroidOrderKey = "lB": Exit Function
    Case DORDER_DEMOLISH: getDroidOrderKey = "D": Exit Function
    Case DORDER_REPAIR: getDroidOrderKey = "R": Exit Function
    Case DORDER_OBSERVE: getDroidOrderKey = "O": Exit Function
    Case DORDER_FIRESUPPORT: getDroidOrderKey = "F": Exit Function
    Case DORDER_UNUSED_4: getDroidOrderKey = "Err": Exit Function
    Case DORDER_UNUSED_2: getDroidOrderKey = "Err": Exit Function
    Case DORDER_RTB: getDroidOrderKey = "RTB": Exit Function
    Case DORDER_RTR: getDroidOrderKey = "RTR": Exit Function
    Case DORDER_UNUSED_5: getDroidOrderKey = "Err": Exit Function
    Case DORDER_EMBARK: getDroidOrderKey = "E": Exit Function
    Case DORDER_DISEMBARK: getDroidOrderKey = "!E": Exit Function
    Case DORDER_ATTACKTARGET: getDroidOrderKey = "AT": Exit Function
    Case DORDER_COMMANDERSUPPORT: getDroidOrderKey = "CS": Exit Function
    Case DORDER_BUILDMODULE: getDroidOrderKey = "BM": Exit Function
    Case DORDER_RECYCLE: getDroidOrderKey = "RCY": Exit Function
    Case DORDER_TRANSPORTOUT: getDroidOrderKey = "To": Exit Function
    Case DORDER_TRANSPORTIN: getDroidOrderKey = "Ti": Exit Function
    Case DORDER_TRANSPORTRETURN: getDroidOrderKey = "Tr": Exit Function
    Case DORDER_GUARD: getDroidOrderKey = "G": Exit Function
    Case DORDER_DROIDREPAIR: getDroidOrderKey = "DR": Exit Function
    Case DORDER_RESTORE: getDroidOrderKey = "RES": Exit Function
    Case DORDER_SCOUT: getDroidOrderKey = "S": Exit Function
    Case DORDER_UNUSED_3: getDroidOrderKey = "Err": Exit Function
    Case DORDER_UNUSED: getDroidOrderKey = "Err": Exit Function
    Case DORDER_PATROL: getDroidOrderKey = "P": Exit Function
    Case DORDER_REARM: getDroidOrderKey = "RE": Exit Function
    Case DORDER_RECOVER: getDroidOrderKey = "RCV": Exit Function
    Case DORDER_UNUSED_6: getDroidOrderKey = "Err": Exit Function
    Case DORDER_RTR_SPECIFIED: getDroidOrderKey = "RTR": Exit Function
    Case DORDER_CIRCLE: getDroidOrderKey = "C": Exit Function
    Case DORDER_HOLD: getDroidOrderKey = "H": Exit Function
    End Select
    ASSERT_FAILURE "false", "src/order.cpp:4295", "getDroidOrderKey", "DROID_ORDER out of range: %u", order
    getDroidOrderKey = "Err"
End Function
