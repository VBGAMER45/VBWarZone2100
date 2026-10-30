Attribute VB_Name = "ai_c"
' src/ai.cpp, ai.h: the AI system (the alliances, the targeting of droids
' and structures, aiUpdateDroid). The AllianceType constants are in
' stubs_game (with configuration.cpp's users); ALLIANCES is ALLIANCES_
' there (VB would take it for the array alliances).
' (VB) A TARGET_ORIGIN *targetOrigin that may be nullptr (aiChooseTarget's)
' is the Optional pair hasTargetOrigin, targetOrigin: hasTargetOrigin
' False is C++'s nullptr. aiSearchSensorTargets' callers always pass one:
' a plain ByRef. A BASE_OBJECT **ppsObj is a ByRef handle.
Option Explicit

Public Const ALLIANCE_BROKEN As Long = 0    ' states of alliance between players
Public Const ALLIANCE_REQUESTED As Long = 1
Public Const ALLIANCE_INVITATION As Long = 2
Public Const ALLIANCE_FORMED As Long = 3
Public Const ALLIANCE_NULL As Long = 4      ' for setting values only.

' Amount of time to rage at the world when frustrated (10 seconds)
Public Const FRUSTRATED_TIME As Long = 1000 * 10

' alliances
Public alliances(0 To MAX_PLAYER_SLOTS - 1, 0 To MAX_PLAYER_SLOTS - 1) As Long ' uint8_t
Public alliancebits(0 To MAX_PLAYER_SLOTS - 1) As Long                         ' PlayerMask
Public satuplinkbits As Long                                                    ' PlayerMask

' Weights used for target selection code,
' target distance is used as 'common currency'
Private Const WEIGHT_DIST_TILE As Long = 13                                    ' In points used in weaponmodifier.txt and structuremodifier.txt
Private Const WEIGHT_DIST_TILE_DROID As Long = WEIGHT_DIST_TILE                ' How much weight a distance of 1 tile (128 world units) has when looking for the best nearest target
Private Const WEIGHT_DIST_TILE_STRUCT As Long = WEIGHT_DIST_TILE
Private Const WEIGHT_HEALTH_DROID As Long = WEIGHT_DIST_TILE * 10              ' How much weight unit damage has (100% of damage is equally weighted as 10 tiles distance)
'~100% damage should be ~8 tiles (max sensor range)
Private Const WEIGHT_HEALTH_STRUCT As Long = WEIGHT_DIST_TILE * 7

Private Const WEIGHT_NOT_VISIBLE_F As Long = 10                                ' We really don't like objects we can't see
Private Const WEIGHT_NOT_LOS_VISIBLE_F As Long = 4

Private Const WEIGHT_SERVICE_DROIDS As Long = WEIGHT_DIST_TILE_DROID * 5       ' We don't want them to be repairing droids or structures while we are after them
Private Const WEIGHT_WEAPON_DROIDS As Long = WEIGHT_DIST_TILE_DROID * 4        ' We prefer to go after anything that has a gun and can hurt us
Private Const WEIGHT_COMMAND_DROIDS As Long = WEIGHT_DIST_TILE_DROID * 6       ' Commanders get a higher priority
Private Const WEIGHT_MILITARY_STRUCT As Long = WEIGHT_DIST_TILE_STRUCT          ' Droid/cyborg factories, repair facility; shouldn't have too much weight
Private Const WEIGHT_WEAPON_STRUCT As Long = WEIGHT_WEAPON_DROIDS              ' Same as weapon droids (?)
Private Const WEIGHT_DERRICK_STRUCT As Long = WEIGHT_MILITARY_STRUCT + WEIGHT_DIST_TILE_STRUCT * 4  ' Even if it's 4 tiles further away than defenses we still choose it

Private Const WEIGHT_STRUCT_NOTBUILT_F As Long = 8                             ' Humans won't fool us anymore!

Private Const OLD_TARGET_THRESHOLD As Long = WEIGHT_DIST_TILE * 4              ' it only makes sense to switch target if new one is 4+ tiles closer

Private Const EMP_DISABLED_PENALTY_F As Long = 10                              ' EMP shouldn't attack emped targets again
Private Const EMP_STRUCT_PENALTY_F As Long = EMP_DISABLED_PENALTY_F * 2        ' EMP don't attack structures, should be bigger than EMP_DISABLED_PENALTY_F

Private Const TOO_CLOSE_PENALTY_F As Long = 20

Private Const TARGET_DOOMED_PENALTY_F As Long = 10      ' Targets that have a lot of damage incoming are less attractive
Private Const TARGET_DOOMED_SLOW_RELOAD_T As Long = 21  ' Weapon ROF threshold for above penalty. per minute.

'Some weights for the units attached to a commander
Private Const WEIGHT_CMD_RANK As Long = WEIGHT_DIST_TILE * 4                   ' A single rank is as important as 4 tiles distance
Private Const WEIGHT_CMD_SAME_TARGET As Long = WEIGHT_DIST_TILE                ' Don't want this to be too high, since a commander can have many units assigned

Private numDroidNearestTargetChecksThisFrame As Long    ' size_t
Private lastGameTimeCheckedNeartestTargets As Long      ' UDWORD

Private gridListNearest() As Long       ' aiBestNearestTarget's static GridList
Private gridListChoose() As Long        ' aiChooseTarget's static GridList
Private gridListSensor() As Long        ' aiChooseSensorTarget's static GridList

' Check no alliance has formed. This is a define to make sure we inline it.
Public Function aiCheckAlliances(ByVal s1 As Long, ByVal s2 As Long) As Boolean
    aiCheckAlliances = alliances(s1, s2) = ALLIANCE_FORMED
End Function

' (VB) a BASE_OBJECT's numWeaps
Private Function objNumWeaps(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: objNumWeaps = aDroid(psObj).numWeaps
    Case OBJ_STRUCTURE: objNumWeaps = aStruct(psObj).numWeaps
    Case OBJ_FEATURE: objNumWeaps = aFeature(psObj).numWeaps
    End Select
End Function

Private Function aiDroidRange(ByVal psDroid As Long, ByVal weapon_slot As Long) As Long
    Dim longRange As Long, psWStats As Long
    If aDroid(psDroid).droidType = DROID_SENSOR Then
        longRange = objSensorRange(psDroid)
    ElseIf aDroid(psDroid).numWeaps = 0 Or aDroid(psDroid).asWeaps(0).nStat = 0 Then
        ' Can't attack without a weapon
        aiDroidRange = 0
        Exit Function
    Else
        psWStats = DROID_getWeaponStats(psDroid, weapon_slot)
        longRange = proj_GetLongRange(asWeaponStats(psWStats), aDroid(psDroid).player)
    End If

    aiDroidRange = longRange
End Function

' see if a structure has the range to fire on a target
Private Function aiStructHasRange(ByVal psStruct As Long, ByVal psTarget As Long, ByVal weapon_slot As Long) As Boolean
    Dim psWStats As Long, longRange As Long
    If aStruct(psStruct).numWeaps = 0 Or aStruct(psStruct).asWeaps(0).nStat = 0 Then
        ' Can't attack without a weapon
        aiStructHasRange = False
        Exit Function
    End If

    psWStats = BASE_OBJECT_getWeaponStats(psStruct, weapon_slot)

    longRange = proj_GetLongRange(asWeaponStats(psWStats), aStruct(psStruct).player)
    If objPosDiffSq_obj(psStruct, psTarget) < longRange * longRange Then
        aiStructHasRange = lineOfFire(psStruct, psTarget, weapon_slot, True)
    End If
End Function

Private Function aiDroidHasRange(ByVal psDroid As Long, ByVal psTarget As Long, ByVal weapon_slot As Long) As Boolean
    Dim longRange As Long
    longRange = aiDroidRange(psDroid, weapon_slot)

    aiDroidHasRange = objPosDiffSq_obj(psDroid, psTarget) < longRange * longRange
End Function

Private Function aiObjHasRange(ByVal psObj As Long, ByVal psTarget As Long, ByVal weapon_slot As Long) As Boolean
    If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        aiObjHasRange = aiDroidHasRange(psObj, psTarget, weapon_slot)
        Exit Function
    ElseIf SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE Then
        aiObjHasRange = aiStructHasRange(psObj, psTarget, weapon_slot)
        Exit Function
    End If
    aiObjHasRange = False
End Function

' Initialise the AI system
Public Function aiInitialise() As Boolean
    Dim i As Long, j As Long, valid As Boolean

    For i = 0 To MAX_PLAYER_SLOTS - 1
        alliancebits(i) = 0
        For j = 0 To MAX_PLAYER_SLOTS - 1
            valid = (i = j And i < MAX_PLAYERS)

            If valid Then alliances(i, j) = ALLIANCE_FORMED Else alliances(i, j) = ALLIANCE_BROKEN
            If valid Then alliancebits(i) = alliancebits(i) Or BIT_(j)
        Next
    Next
    satuplinkbits = 0

    aiInitialise = True
End Function

' Check properties of the AllianceType enum.
Public Function alliancesFixed(ByVal t As Long) As Boolean
    alliancesFixed = t <> ALLIANCES_
End Function

Public Function alliancesSharedVision(ByVal t As Long) As Boolean
    alliancesSharedVision = t = ALLIANCES_TEAMS Or t = ALLIANCES_UNSHARED
End Function

Public Function alliancesSharedResearch(ByVal t As Long) As Boolean
    alliancesSharedResearch = t = ALLIANCES_ Or t = ALLIANCES_TEAMS
End Function

Public Function alliancesSetTeamsBeforeGame(ByVal t As Long) As Boolean
    alliancesSetTeamsBeforeGame = t = ALLIANCES_TEAMS Or t = ALLIANCES_UNSHARED
End Function

Public Function alliancesCanGiveResearchAndRadar(ByVal t As Long) As Boolean
    alliancesCanGiveResearchAndRadar = t = ALLIANCES_
End Function
' Shutdown the AI system
Public Function aiShutdown() As Boolean
    aiShutdown = True
End Function

' Search the global list of sensors for a possible target for psObj.
Private Function aiSearchSensorTargets(ByVal psObj As Long, ByVal weapon_slot As Long, ByVal psWStats As Long, targetOrigin As TARGET_ORIGIN) As Long
    Dim longRange As Long, tarDist As Long, foundCB As Boolean, minDist As Long, psTarget As Long
    Dim it As Long, psSensor As Long, psTemp As Long, isCB As Boolean, psDroid As Long, sensorType As Long
    Dim psCStruct As Long, distSq As Long, objPlayer As Long, c As Boolean

    objPlayer = SIMPLE_OBJECT_player(psObj)
    longRange = proj_GetLongRange(asWeaponStats(psWStats), objPlayer)
    tarDist = longRange * longRange
    foundCB = False
    minDist = proj_GetMinRange(asWeaponStats(psWStats), objPlayer) * proj_GetMinRange(asWeaponStats(psWStats), objPlayer)
    psTarget = 0

    targetOrigin = ORIGIN_UNKNOWN

    it = apsSensorList(0).head
    Do While it <> 0
        psSensor = lnObj(it)
        psTemp = 0
        isCB = False
        'bool isRD = false;

        If Not aiCheckAlliances(SIMPLE_OBJECT_player(psSensor), objPlayer) Then
            GoTo continue_
        ElseIf SIMPLE_OBJECT_type(psSensor) = OBJ_DROID Then
            psDroid = psSensor

            If Not (aDroid(psDroid).droidType = DROID_SENSOR) Then ASSERT_FAILURE "psDroid->droidType == DROID_SENSOR", "src/ai.cpp:196", "aiSearchSensorTargets", "A non-sensor droid in a sensor list is non-sense": aiSearchSensorTargets = 0: Exit Function
            ' Skip non-observing droids. This includes Radar Detectors at the moment since they never observe anything.
            If aDroid(psDroid).action <> DACTION_OBSERVE Then
                GoTo continue_
            End If
            sensorType = asSensorStats(DROID_getSensorStats(psDroid)).typ
            ' Artillery should not fire at objects observed by VTOL CB/Strike sensors.
            If sensorType = VTOL_CB_SENSOR Or _
               sensorType = VTOL_INTERCEPT_SENSOR Then
                GoTo continue_
            ElseIf objRadarDetector(psDroid) Then
                GoTo continue_
            End If
            psTemp = aDroid(psDroid).psActionTarget(0)
            isCB = sensorType = INDIRECT_CB_SENSOR
            'isRD = objRadarDetector((BASE_OBJECT *)psDroid);
        ElseIf SIMPLE_OBJECT_type(psSensor) = OBJ_STRUCTURE Then
            psCStruct = psSensor

            ' skip incomplete structures
            If aStruct(psCStruct).status <> SS_BUILT Then
                GoTo continue_
            End If
            ' Artillery should not fire at objects observed by VTOL CB/Strike sensors.
            sensorType = asSensorStats(asStructureStats(aStruct(psCStruct).pStructureType).pSensor).typ
            If sensorType = VTOL_CB_SENSOR Or _
               sensorType = VTOL_INTERCEPT_SENSOR Then
                GoTo continue_
            ElseIf objRadarDetector(psCStruct) Then
                GoTo continue_
            End If
            psTemp = aStruct(psCStruct).psTarget(0)
            isCB = structCBSensor(psCStruct)
            'isRD = objRadarDetector((BASE_OBJECT *)psCStruct);
        End If
        If psTemp = 0 Then GoTo continue_
        If SIMPLE_OBJECT_died(psTemp) <> 0 Then GoTo continue_
        If aiObjectIsProbablyDoomed(psTemp, False) Then GoTo continue_
        If Not validTarget(psObj, psTemp, 0) Then GoTo continue_
        If aiCheckAlliances(SIMPLE_OBJECT_player(psTemp), objPlayer) Then GoTo continue_
        distSq = objPosDiffSq_obj(psTemp, psObj)
        ' Need to be in range, prefer closer targets or CB targets
        c = (isCB And Not foundCB)
        If Not c Then c = (isCB = foundCB And distSq < tarDist)
        If c And distSq > minDist Then
            c = aiObjHasRange(psObj, psTemp, weapon_slot)
            If c Then c = visibleObject(psSensor, psTemp, False) <> 0
            If c Then
                tarDist = distSq
                psTarget = psTemp
                targetOrigin = ORIGIN_SENSOR

                If isCB Then
                    targetOrigin = ORIGIN_CB_SENSOR
                    foundCB = True  ' got CB target, drop everything and shoot!
                End If
                '
                'else if (isRD)
                '{
                '    if (targetOrigin)
                '    {
                '        *targetOrigin = ORIGIN_RADAR_DETECTOR;
                '    }
                '}
                '
            End If
        End If
continue_:
        it = lnNext(it)
    Loop
    aiSearchSensorTargets = psTarget
End Function

' Calculates attack priority for a certain target
' Returns the target attack weight if it's greater than currentBest
' (otherwise, may shortcut various costly calculations, and returns <= 0)
Private Function targetAttackWeightIfGreaterThan(ByVal currentBest As Long, ByVal psTarget As Long, ByVal psAttacker As Long, ByVal weapon_slot As Long) As Long
    Dim targetTypeBonus As Long, damageRatio As Long, attackWeight As Long, noTarget As Long, weaponSlot As Long
    Dim targetDroid As Long, psAttackerDroid As Long, psDroid As Long, targetStructure As Long, weaponEffect As WEAPON_EFFECT
    Dim attackerWeapon As Long, bEmpWeap As Boolean, bCmdAttached As Boolean, bTargetingCmd As Boolean, bDirect As Boolean
    Dim dist As Long, tooClose As Boolean, apos As Vector3i, tpos As Vector3i, commander As Long, it As Long, psGroupDroid As Long

    targetTypeBonus = 0
    damageRatio = 0
    attackWeight = 0
    noTarget = -1
    targetDroid = 0
    psAttackerDroid = 0
    targetStructure = 0
    bEmpWeap = False
    bCmdAttached = False
    bTargetingCmd = False
    bDirect = False

    If psTarget = 0 Or psAttacker = 0 Then
        targetAttackWeightIfGreaterThan = noTarget
        Exit Function
    ElseIf SIMPLE_OBJECT_died(psTarget) <> 0 Then
        targetAttackWeightIfGreaterThan = noTarget
        Exit Function
    End If
    If Not (psTarget <> psAttacker) Then ASSERT_FAILURE "psTarget != psAttacker", "src/ai.cpp:291", "targetAttackWeightIfGreaterThan", "targetAttackWeight: Wanted to evaluate the worth of attacking ourselves..."

    targetTypeBonus = 0 ' Sensors/ecm droids, non-military structures get lower priority

    ' Get attacker weapon effect
    If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
        psAttackerDroid = psAttacker

        attackerWeapon = DROID_getWeaponStats(psAttackerDroid, weapon_slot)

        'check if this droid is assigned to a commander
        bCmdAttached = hasCommander(psAttackerDroid)

        'find out if current target is targeting our commander
        If bCmdAttached Then
            commander = aGroup(aDroid(psAttackerDroid).psGroup).psCommander
            If SIMPLE_OBJECT_type(psTarget) = OBJ_DROID Then
                psDroid = psTarget

                'go through all enemy weapon slots
                weaponSlot = 0
                Do While Not bTargetingCmd And ULT(weaponSlot, aDroid(psTarget).numWeaps)
                    'see if this weapon is targeting our commander
                    If aDroid(psDroid).psActionTarget(weaponSlot) = commander Then
                        bTargetingCmd = True
                    End If
                    weaponSlot = weaponSlot + 1
                Loop
            Else
                If SIMPLE_OBJECT_type(psTarget) = OBJ_STRUCTURE Then
                    'go through all enemy weapons
                    weaponSlot = 0
                    Do While Not bTargetingCmd And ULT(weaponSlot, aStruct(psTarget).numWeaps)
                        If aStruct(psTarget).psTarget(weaponSlot) = _
                           commander Then
                            bTargetingCmd = True
                        End If
                        weaponSlot = weaponSlot + 1
                    Loop
                End If
            End If
        End If
    ElseIf SIMPLE_OBJECT_type(psAttacker) = OBJ_STRUCTURE Then
        attackerWeapon = BASE_OBJECT_getWeaponStats(psAttacker, weapon_slot)
    Else    ' feature
        ASSERT_FAILURE "!""invalid attacker object type""", "src/ai.cpp:346", "targetAttackWeightIfGreaterThan", "targetAttackWeight: Invalid attacker object type"
        targetAttackWeightIfGreaterThan = noTarget
        Exit Function
    End If

    bDirect = proj_Direct(attackerWeapon)
    If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
        If aDroid(psAttackerDroid).droidType = DROID_SENSOR Then
            ' Sensors are considered a direct weapon,
            ' but for computing expected damage it makes more sense to use indirect damage
            bDirect = False
        End If
    End If

    'Get weapon effect
    weaponEffect = asWeaponStats(attackerWeapon).weaponEffect

    'See if attacker is using an EMP weapon
    bEmpWeap = (asWeaponStats(attackerWeapon).weaponSubClass = WSC_EMP)

    apos = SIMPLE_OBJECT_pos(psAttacker)
    tpos = SIMPLE_OBJECT_pos(psTarget)
    dist = iHypot(apos.x - tpos.x, apos.y - tpos.y)
    tooClose = Not ULT(proj_GetMinRange(asWeaponStats(attackerWeapon), SIMPLE_OBJECT_player(psAttacker)), dist)
    If tooClose Then
        dist = objSensorRange(psAttacker)  ' If object is too close to fire at, consider it to be at maximum range.
    End If

    ' Calculate attack weight
    If SIMPLE_OBJECT_type(psTarget) = OBJ_DROID Then
        targetDroid = psTarget

        If aDroid(targetDroid).died <> 0 Then
            debug_ 378, LOG_NEVER, "targetAttackWeightIfGreaterThan", "Target droid is dead, skipping invalid droid.\n"
            targetAttackWeightIfGreaterThan = noTarget
            Exit Function
        End If

        ' Calculate damage this target suffered
        If aDroid(targetDroid).originalBody = 0 Then ' FIXME Somewhere we get 0HP droids from
            damageRatio = 0
            debug_ 386, LOG_ERROR, "targetAttackWeightIfGreaterThan", "targetAttackWeight: 0HP droid detected!"
            debug_ 387, LOG_ERROR, "targetAttackWeightIfGreaterThan", "  Type: %i Name: ""%s"" Owner: %i ""%s"")", _
                   aDroid(targetDroid).droidType, aDroid(targetDroid).aName, aDroid(targetDroid).player, getPlayerName(aDroid(targetDroid).player)
        Else
            damageRatio = FromU32(100 - Fix(U32(MUL32(100, aDroid(targetDroid).body)) / U32(aDroid(targetDroid).originalBody)))
        End If
        ' assert(targetDroid->originalBody != 0): a release build's assert (none)

        ' See if this type of a droid should be prioritized
        Select Case aDroid(targetDroid).droidType
        Case DROID_SENSOR, DROID_ECM, DROID_PERSON, DROID_TRANSPORTER, DROID_SUPERTRANSPORTER, DROID_DEFAULT, DROID_ANY

        Case DROID_CYBORG, DROID_WEAPON, DROID_CYBORG_SUPER
            targetTypeBonus = WEIGHT_WEAPON_DROIDS

        Case DROID_COMMAND
            targetTypeBonus = WEIGHT_COMMAND_DROIDS

        Case DROID_CONSTRUCT, DROID_REPAIR, DROID_CYBORG_CONSTRUCT, DROID_CYBORG_REPAIR
            targetTypeBonus = WEIGHT_SERVICE_DROIDS
        End Select

        ' Now calculate the overall weight
        attackWeight = asWeaponModifier(weaponEffect, asPropulsionStats(DROID_getPropulsionStats(targetDroid)).propulsionType) _
                       + asWeaponModifierBody(weaponEffect, asBodyStats(DROID_getBodyStats(targetDroid)).size) _
                       + WEIGHT_DIST_TILE_DROID * objSensorRange(psAttacker) \ TILE_UNITS _
                       - WEIGHT_DIST_TILE_DROID * dist \ TILE_UNITS _
                       + WEIGHT_HEALTH_DROID * damageRatio \ 100 _
                       + targetTypeBonus
        ' Our weapon's effect against target; farther droids are less attractive; we prefer damaged droids;
        ' some droid types have higher priority

        ' If attacking with EMP try to avoid targets that were already "EMPed"
        If bEmpWeap Then
            If aDroid(targetDroid).lastHitWeapon = WSC_EMP Then
                If ULT(FromU32(U32(gameTime) - U32(aDroid(targetDroid).timeLastHit)), EMP_DISABLE_TIME) Then  ' target still disabled
                    attackWeight = attackWeight \ EMP_DISABLED_PENALTY_F
                End If
            End If
        End If
    ElseIf SIMPLE_OBJECT_type(psTarget) = OBJ_STRUCTURE Then
        targetStructure = psTarget

        ' Calculate damage this target suffered
        damageRatio = FromU32(100 - Fix(U32(MUL32(100, aStruct(targetStructure).body)) / U32(STRUCTURE_structureBody(targetStructure))))

        ' See if this type of a structure should be prioritized
        Select Case asStructureStats(aStruct(targetStructure).pStructureType).typ
        Case REF_DEFENSE, REF_FORTRESS
            targetTypeBonus = WEIGHT_WEAPON_STRUCT

        Case REF_RESOURCE_EXTRACTOR
            targetTypeBonus = WEIGHT_DERRICK_STRUCT

        Case REF_FACTORY, REF_CYBORG_FACTORY, REF_REPAIR_FACILITY
            targetTypeBonus = WEIGHT_MILITARY_STRUCT
        Case Else
        End Select

        ' Now calculate the overall weight
        attackWeight = asStructStrengthModifier(weaponEffect, asStructureStats(aStruct(targetStructure).pStructureType).strength) _
                       + WEIGHT_DIST_TILE_STRUCT * objSensorRange(psAttacker) \ TILE_UNITS _
                       - WEIGHT_DIST_TILE_STRUCT * dist \ TILE_UNITS _
                       + WEIGHT_HEALTH_STRUCT * damageRatio \ 100 _
                       + targetTypeBonus
        ' Our weapon's effect against target; farther structs are less attractive; we prefer damaged structures;
        ' some structure types have higher priority

        ' Go for unfinished structures only if nothing else found (same for non-visible structures)
        If aStruct(targetStructure).status <> SS_BUILT Then   'a decoy?
            attackWeight = attackWeight \ WEIGHT_STRUCT_NOTBUILT_F
        End If

        ' EMP should only attack structures if no enemy droids are around
        If bEmpWeap Then
            attackWeight = attackWeight \ EMP_STRUCT_PENALTY_F
        End If
    Else    'a feature
        targetAttackWeightIfGreaterThan = 1
        Exit Function
    End If

    ' We prefer objects we can see and can attack immediately
    If visibleObject(psAttacker, psTarget, True) = 0 Then
        attackWeight = attackWeight \ WEIGHT_NOT_VISIBLE_F
    End If

    If tooClose Then
        attackWeight = attackWeight \ TOO_CLOSE_PENALTY_F
    End If

    ' Penalty for units that are already considered doomed (but the missile might miss!)
    If aiObjectIsProbablyDoomed(psTarget, bDirect) Then
        ' indirect firing units have slow reload times, so give the target a chance to die,
        ' and give a different unit a chance to get in range, too.
        If weaponROF(asWeaponStats(attackerWeapon), SIMPLE_OBJECT_player(psAttacker)) < TARGET_DOOMED_SLOW_RELOAD_T Then
            debug_ 512, LOG_NEVER, "targetAttackWeightIfGreaterThan", "Not killing unit - doomed. My ROF: %i (%s)", weaponROF(asWeaponStats(attackerWeapon), SIMPLE_OBJECT_player(psAttacker)), asWeaponStats(attackerWeapon).name_
            targetAttackWeightIfGreaterThan = noTarget
            Exit Function
        End If
        attackWeight = attackWeight \ TARGET_DOOMED_PENALTY_F
    End If

    ' Commander-related criterias
    If bCmdAttached Then    'attached to a commander and don't have a target assigned by some order
        commander = aGroup(aDroid(psAttackerDroid).psGroup).psCommander
        If Not (commander <> 0) Then ASSERT_FAILURE "psAttackerDroid->psGroup->psCommander != nullptr", "src/ai.cpp:521", "targetAttackWeightIfGreaterThan", "Commander is NULL"

        'if commander is being targeted by our target, try to defend the commander
        If bTargetingCmd Then
            attackWeight = attackWeight + WEIGHT_CMD_RANK * (1 + getDroidLevel(commander))
        End If

        'fire support - go through all droids assigned to the commander
        it = aGroup(aDroid(psAttackerDroid).psGroup).psList.head
        Do While it <> 0
            psGroupDroid = lnObj(it)
            For weaponSlot = 0 To aDroid(psGroupDroid).numWeaps - 1
                'see if this droid is currently targeting current target
                If aDroid(psGroupDroid).order.psObj = psTarget Or _
                   aDroid(psGroupDroid).psActionTarget(weaponSlot) = psTarget Then
                    'we prefer targets that are already targeted and hence will be destroyed faster
                    attackWeight = attackWeight + WEIGHT_CMD_SAME_TARGET
                End If
            Next
            it = lnNext(it)
        Loop
    End If

    If bDirect Then
        ' lineOfFire is a costly check - shortcut if there's no way we'd return an attackWeight > currentBest
        If attackWeight <= currentBest Then
            targetAttackWeightIfGreaterThan = MIN_(0, attackWeight)
            Exit Function
        End If
        If Not lineOfFire(psAttacker, psTarget, weapon_slot, False) Then
            attackWeight = attackWeight \ WEIGHT_NOT_LOS_VISIBLE_F ' Prefer objects not obstructed by terrain
        End If
    End If

    targetAttackWeightIfGreaterThan = MAX_(1, attackWeight)
End Function

Public Function getCountNearestTargetChecks() As Long
    getCountNearestTargetChecks = numDroidNearestTargetChecksThisFrame
End Function

' Find the best nearest target for a droid.
' If extraRange is higher than zero, then this is the range it accepts for movement to target.
' Returns integer representing target priority, -1 if failed
Public Function aiBestNearestTarget(ByVal psDroid As Long, ppsObj As Long, ByVal weapon_slot As Long, Optional ByVal extraRange As Long = 0) As Long
    Dim failure As Long, bestMod As Long, psTarget As Long, bestTarget As Long, tempTarget As Long, electronic As Boolean
    Dim targetStructure As Long, weaponEffect As WEAPON_EFFECT, tmpOrigin As TARGET_ORIGIN, psWStats As Long, droidRange As Long
    Dim n As Long, gi As Long, friendlyObj As Long, targetInQuestion As Long, friendlyDroid As Long, psStruct As Long
    Dim tqType As OBJECT_TYPE, newMod As Long, c As Boolean

    failure = -1
    bestMod = 0
    psTarget = 0
    bestTarget = 0
    electronic = False
    tmpOrigin = ORIGIN_UNKNOWN

    'don't bother looking if empty vtol droid
    If vtolEmpty(psDroid) Then
        aiBestNearestTarget = failure
        Exit Function
    End If

    ' Return if have no weapons
    ' The ai orders a non-combat droid to patrol = crash without it...
    If (aDroid(psDroid).asWeaps(0).nStat = 0 Or aDroid(psDroid).numWeaps = 0) And aDroid(psDroid).droidType <> DROID_SENSOR Then
        aiBestNearestTarget = failure
        Exit Function
    End If

    If lastGameTimeCheckedNeartestTargets <> gameTime Then
        lastGameTimeCheckedNeartestTargets = gameTime
        numDroidNearestTargetChecksThisFrame = 0
    End If
    numDroidNearestTargetChecksThisFrame = numDroidNearestTargetChecksThisFrame + 1

    ' Check if we have a CB target to begin with
    psWStats = DROID_getWeaponStats(psDroid, weapon_slot)
    If Not proj_Direct(psWStats) Then
        bestTarget = aiSearchSensorTargets(psDroid, weapon_slot, psWStats, tmpOrigin)
        bestMod = targetAttackWeightIfGreaterThan(-1, bestTarget, psDroid, weapon_slot)
    End If

    weaponEffect = asWeaponStats(psWStats).weaponEffect

    electronic = electronicDroid(psDroid)

    ' Range was previously 9*TILE_UNITS. Increasing this doesn't seem to help much, though. Not sure why.
    droidRange = MIN_(aiDroidRange(psDroid, weapon_slot) + extraRange, objSensorRange(psDroid) + 6 * TILE_UNITS)

    n = gridStartIterate(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, droidRange, gridListNearest)
    For gi = 0 To n - 1
        friendlyObj = 0
        targetInQuestion = gridListNearest(gi)

        If targetInQuestion = 0 Then
            GoTo continue_
        ElseIf isDead(targetInQuestion) Then
            GoTo continue_
        End If

        ' This is a friendly unit, check if we can reuse its target
        If aiCheckAlliances(SIMPLE_OBJECT_player(targetInQuestion), aDroid(psDroid).player) Then
            friendlyObj = targetInQuestion
            targetInQuestion = 0

            ' Can we see what it is doing?
            If BASE_OBJECT_visible(friendlyObj, aDroid(psDroid).player) = UBYTE_MAX Then
                If SIMPLE_OBJECT_type(friendlyObj) = OBJ_DROID Then
                    friendlyDroid = friendlyObj

                    ' See if friendly droid has a target
                    tempTarget = aDroid(friendlyDroid).psActionTarget(0)
                    If tempTarget <> 0 Then
                        If SIMPLE_OBJECT_died(tempTarget) = 0 Then
                            'make sure a weapon droid is targeting it
                            If aDroid(friendlyDroid).numWeaps > 0 Then
                                ' make sure this target wasn't assigned explicitly to this droid
                                If aDroid(friendlyDroid).order.typ <> DORDER_ATTACK Then
                                    If bestTarget <> tempTarget And psTarget <> tempTarget Then
                                        targetInQuestion = tempTarget  'consider this target
                                    End If
                                End If
                            End If
                        End If
                    End If
                ElseIf SIMPLE_OBJECT_type(friendlyObj) = OBJ_STRUCTURE Then
                    tempTarget = aStruct(friendlyObj).psTarget(0)
                    If tempTarget <> 0 Then
                        If SIMPLE_OBJECT_died(tempTarget) = 0 Then
                            If bestTarget <> tempTarget And psTarget <> tempTarget Then
                                targetInQuestion = tempTarget
                            End If
                        End If
                    End If
                End If
            End If
        End If

        c = False
        If targetInQuestion <> 0 Then
            If targetInQuestion <> psDroid Then  ' in case friendly unit had me as target
                tqType = SIMPLE_OBJECT_type(targetInQuestion)
                If tqType = OBJ_DROID Or tqType = OBJ_STRUCTURE Or tqType = OBJ_FEATURE Then
                    If BASE_OBJECT_visible(targetInQuestion, aDroid(psDroid).player) = UBYTE_MAX Then
                        If Not aiCheckAlliances(SIMPLE_OBJECT_player(targetInQuestion), aDroid(psDroid).player) Then
                            If validTarget(psDroid, targetInQuestion, weapon_slot) Then
                                c = objPosDiffSq_obj(psDroid, targetInQuestion) < droidRange * droidRange
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If c Then
            If tqType = OBJ_DROID Then
                ' in multiPlayer - don't attack Transporters with EW
                If bMultiPlayer Then
                    ' if not electronic then valid target
                    If Not electronic Then
                        'only a valid target if NOT a transporter
                        psTarget = targetInQuestion
                    ElseIf Not DROID_isTransporter(targetInQuestion) Then
                        psTarget = targetInQuestion
                    End If
                Else
                    psTarget = targetInQuestion
                End If
            ElseIf tqType = OBJ_STRUCTURE Then
                psStruct = targetInQuestion

                If electronic Then
                    ' don't want to target structures with resistance of zero if using electronic warfare
                    If validStructResistance(targetInQuestion) Then
                        psTarget = targetInQuestion
                    End If
                ElseIf aStruct(psStruct).asWeaps(0).nStat > 0 Then
                    ' structure with weapons - go for this
                    psTarget = targetInQuestion
                ElseIf isHumanPlayer(aDroid(psDroid).player) And (asStructureStats(aStruct(psStruct).pStructureType).typ <> REF_WALL And asStructureStats(aStruct(psStruct).pStructureType).typ <> REF_WALLCORNER) Then
                    psTarget = targetInQuestion
                ElseIf Not isHumanPlayer(aDroid(psDroid).player) Then
                    psTarget = targetInQuestion
                End If
            ElseIf tqType = OBJ_FEATURE Then
                If aDroid(psDroid).lastFrustratedTime > 0 Then
                    If ULT(FromU32(U32(gameTime) - U32(aDroid(psDroid).lastFrustratedTime)), FRUSTRATED_TIME) Then
                        If asFeatureStats(aFeature(targetInQuestion).psStats).damageable Then
                            If aDroid(psDroid).player <> scavengerPlayer() Then  ' hack to avoid scavs blowing up their nice feature walls
                                psTarget = targetInQuestion
                                objTrace aDroid(psDroid).id, "aiBestNearestTarget", "considering shooting at %s in frustration", objInfo(targetInQuestion)
                            End If
                        End If
                    End If
                End If
            End If

            ' Check if our weapon is most effective against this object
            If psTarget <> 0 And psTarget = targetInQuestion _
               And psTarget <> bestTarget Then  ' was assigned?; avoid duplicate checks
                newMod = targetAttackWeightIfGreaterThan(bestMod, psTarget, psDroid, weapon_slot)

                ' Remember this one if it's our best target so far
                If newMod >= 0 And (newMod > bestMod Or bestTarget = 0) Then
                    bestMod = newMod
                    tmpOrigin = ORIGIN_ALLY
                    bestTarget = psTarget
                End If
            End If
        End If
continue_:
    Next

    If bestTarget <> 0 Then
        If Not (SIMPLE_OBJECT_died(bestTarget) = 0) Then ASSERT_FAILURE "!bestTarget->died", "src/ai.cpp:753", "aiBestNearestTarget", "AI gave us a target that is already dead."
        targetStructure = visGetBlockingWall(psDroid, bestTarget)

        ' See if target is blocked by a wall; only affects direct weapons
        ' Ignore friendly walls here
        If proj_Direct(psWStats) And targetStructure <> 0 Then
            If Not aiCheckAlliances(aDroid(psDroid).player, aStruct(targetStructure).player) Then
                'are we any good against walls?
                If asStructStrengthModifier(weaponEffect, asStructureStats(aStruct(targetStructure).pStructureType).strength) >= MIN_STRUCTURE_BLOCK_STRENGTH Then
                    bestTarget = targetStructure 'attack wall
                End If
            End If
        End If

        ppsObj = bestTarget
        aiBestNearestTarget = bestMod
        Exit Function
    End If

    aDroid(psDroid).lastCheckNearestTargetFailed(weapon_slot) = gameTime
    aiBestNearestTarget = failure
End Function

' ---- B7: the target checks the actions and orders use

' Are there a lot of bullets heading towards the droid?
Private Function aiDroidIsProbablyDoomed(ByVal psDroid As Long, ByVal isDirect As Boolean) As Boolean
    Dim e As Double, b As Double
    b = U32(aDroid(psDroid).body)
    If isDirect Then
        e = U32(aDroid(psDroid).expectedDamageDirect)
    Else
        e = U32(aDroid(psDroid).expectedDamageIndirect)
    End If
    ' Doomed if projectiles will damage 120% of remaining body points.
    If e > b Then
        aiDroidIsProbablyDoomed = e - b > Int(b / 5)
    End If
End Function

' Are there a lot of bullets heading towards the structure?
Private Function aiStructureIsProbablyDoomed(ByVal psStructure As Long) As Boolean
    Dim e As Double, b As Double
    b = U32(aStruct(psStructure).body)
    e = U32(aStruct(psStructure).expectedDamage)
    ' Doomed if projectiles will damage 106.6666666667% of remaining body points.
    If e > b Then
        aiStructureIsProbablyDoomed = e - b > Int(b / 15)
    End If
End Function

' Are there a lot of bullets heading towards the object?
Public Function aiObjectIsProbablyDoomed(ByVal psObject As Long, ByVal isDirect As Boolean) As Boolean
    If SIMPLE_OBJECT_died(psObject) <> 0 Then
        aiObjectIsProbablyDoomed = True    ' Was definitely doomed.
        Exit Function
    End If

    Select Case SIMPLE_OBJECT_type(psObject)
    Case OBJ_DROID
        aiObjectIsProbablyDoomed = aiDroidIsProbablyDoomed(psObject, isDirect)
    Case OBJ_STRUCTURE
        aiObjectIsProbablyDoomed = aiStructureIsProbablyDoomed(psObject)
    Case Else
        aiObjectIsProbablyDoomed = False
    End Select
End Function

' Update the expected damage of the object.
Public Sub aiObjectAddExpectedDamage(ByVal psObject As Long, ByVal damage As Long, ByVal isDirect As Boolean)
    If psObject = 0 Then
        Exit Sub    ' Hard to destroy the ground.
    End If

    Select Case SIMPLE_OBJECT_type(psObject)
    Case OBJ_DROID
        If isDirect Then
            aDroid(psObject).expectedDamageDirect = FromU32(CDbl(aDroid(psObject).expectedDamageDirect) + damage)
            If Not (aDroid(psObject).expectedDamageDirect >= 0) Then ASSERT_FAILURE "(SDWORD)((DROID *)psObject)->expectedDamageDirect >= 0", "src/ai.cpp:832", "aiObjectAddExpectedDamage", "aiObjectAddExpectedDamage: Negative amount of projectiles heading towards droid."
        Else
            aDroid(psObject).expectedDamageIndirect = FromU32(CDbl(aDroid(psObject).expectedDamageIndirect) + damage)
            If Not (aDroid(psObject).expectedDamageIndirect >= 0) Then ASSERT_FAILURE "(SDWORD)((DROID *)psObject)->expectedDamageIndirect >= 0", "src/ai.cpp:837", "aiObjectAddExpectedDamage", "aiObjectAddExpectedDamage: Negative amount of projectiles heading towards droid."
        End If
    Case OBJ_STRUCTURE
        aStruct(psObject).expectedDamage = FromU32(CDbl(aStruct(psObject).expectedDamage) + damage)
        If Not (aStruct(psObject).expectedDamage >= 0) Then ASSERT_FAILURE "(SDWORD)((STRUCTURE *)psObject)->expectedDamage >= 0", "src/ai.cpp:842", "aiObjectAddExpectedDamage", "aiObjectAddExpectedDamage: Negative amount of projectiles heading towards droid."
    Case Else
    End Select
End Sub

' see if an object is a wall
Private Function aiObjIsWall(ByVal psObj As Long) As Boolean
    If psObj = 0 Then
        aiObjIsWall = False
        Exit Function
    End If
    If SIMPLE_OBJECT_type(psObj) <> OBJ_STRUCTURE Then
        aiObjIsWall = False
        Exit Function
    End If

    If aStruct(psObj).pStructureType <> -1 Then
        If asStructureStats(aStruct(psObj).pStructureType).typ <> REF_WALL And _
           asStructureStats(aStruct(psObj).pStructureType).typ <> REF_WALLCORNER Then
            aiObjIsWall = False
            Exit Function
        End If
    End If

    aiObjIsWall = True
End Function


' See if there is a target in range
' (VB) hasTargetOrigin False: targetOrigin is nullptr
Public Function aiChooseTarget(ByVal psObj As Long, ppsTarget As Long, ByVal weapon_slot As Long, ByVal bUpdateTarget As Boolean, _
                               Optional ByVal hasTargetOrigin As Boolean = False, Optional targetOrigin As TARGET_ORIGIN = ORIGIN_UNKNOWN) As Boolean
    Dim psTarget As Long, psCommander As Long, curTargetWeight As Long, tmpOrigin As TARGET_ORIGIN, psCurrTarget As Long
    Dim newTargetWeight As Long, c As Boolean, bCommanderBlock As Boolean, psWStats As Long, longRange As Long
    Dim targetValue As Long, tarDist As Long, srange As Long, n As Long, gi As Long, psCurr As Long
    Dim newTargetValue As Long, distSq As Long, objPlayer As Long, objPos As Vector3i

    psTarget = 0
    curTargetWeight = -1
    tmpOrigin = ORIGIN_UNKNOWN

    If hasTargetOrigin Then
        targetOrigin = ORIGIN_UNKNOWN
    End If

    If Not (ULT(weapon_slot, objNumWeaps(psObj))) Then ASSERT_FAILURE "(unsigned)weapon_slot < psObj->numWeaps", "src/ai.cpp:885", "aiChooseTarget", "Invalid weapon selected": aiChooseTarget = False: Exit Function

    objPlayer = SIMPLE_OBJECT_player(psObj)

    ' See if there is a something in range
    If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        psCurrTarget = aDroid(psObj).psActionTarget(0)

        ' find a new target
        newTargetWeight = aiBestNearestTarget(psObj, psTarget, weapon_slot)

        ' Calculate weight of the current target if updating; but take care not to target
        ' ourselves...
        If bUpdateTarget And psCurrTarget <> psObj Then
            curTargetWeight = targetAttackWeightIfGreaterThan(-1, psCurrTarget, psObj, weapon_slot)
        End If

        c = False
        If newTargetWeight >= 0 Then  ' found a new target
            ' choosing a new target, don't care if current one is better; attacker had no valid target, use new one;
            ' updating and new target is better
            If Not bUpdateTarget Or curTargetWeight <= 0 Or newTargetWeight > curTargetWeight + OLD_TARGET_THRESHOLD Then
                If validTarget(psObj, psTarget, weapon_slot) Then
                    c = aiDroidHasRange(psObj, psTarget, weapon_slot)
                End If
            End If
        End If
        If c Then
            If Not (Not isDead(psTarget)) Then ASSERT_FAILURE "!isDead(psTarget)", "src/ai.cpp:909", "aiChooseTarget", "Droid found a dead target!"
            ppsTarget = psTarget
            aiChooseTarget = True
            Exit Function
        End If
    ElseIf SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE Then
        bCommanderBlock = False

        If Not (aStruct(psObj).asWeaps(weapon_slot).nStat > 0) Then ASSERT_FAILURE "psObj->asWeaps[weapon_slot].nStat > 0", "src/ai.cpp:918", "aiChooseTarget", "Invalid weapon turret": aiChooseTarget = False: Exit Function

        psWStats = BASE_OBJECT_getWeaponStats(psObj, weapon_slot)
        longRange = proj_GetLongRange(asWeaponStats(psWStats), objPlayer)

        ' see if there is a target from the command droids
        psTarget = 0
        psCommander = cmdDroidGetDesignator(objPlayer)
        c = False
        If Not proj_Direct(psWStats) And (psCommander <> 0) Then
            c = aiStructHasRange(psObj, psCommander, weapon_slot)
        End If
        If c Then
            ' there is a commander that can fire designate for this structure
            ' set bCommanderBlock so that the structure does not fire until the commander
            ' has a target - (slow firing weapons will not be ready to fire otherwise).
            bCommanderBlock = True

            ' I do believe this will never happen, check for yourself :-)
            debug_ 935, LOG_NEVER, "aiChooseTarget", "Commander %d is good enough for fire designation", aDroid(psCommander).id

            c = False
            If aDroid(psCommander).action = DACTION_ATTACK _
               And aDroid(psCommander).psActionTarget(0) <> 0 Then
                c = SIMPLE_OBJECT_died(aDroid(psCommander).psActionTarget(0)) = 0
            End If
            If c Then
                ' the commander has a target to fire on
                If aiStructHasRange(psObj, aDroid(psCommander).psActionTarget(0), weapon_slot) Then
                    ' target in range - fire on it
                    tmpOrigin = ORIGIN_COMMANDER
                    psTarget = aDroid(psCommander).psActionTarget(0)
                Else
                    ' target out of range - release the commander block
                    bCommanderBlock = False
                End If
            End If
        End If

        ' indirect fire structures use sensor towers first
        If psTarget = 0 And Not bCommanderBlock And Not proj_Direct(psWStats) Then
            psTarget = aiSearchSensorTargets(psObj, weapon_slot, psWStats, tmpOrigin)
        End If

        If psTarget = 0 And Not bCommanderBlock Then
            targetValue = -1
            tarDist = INT32_MAX
            srange = longRange

            If Not proj_Direct(psWStats) And srange > objSensorRange(psObj) Then
                ' search radius of indirect weapons limited by their sight, unless they use
                ' external sensors to provide fire designation
                srange = objSensorRange(psObj)
            End If

            objPos = SIMPLE_OBJECT_pos(psObj)
            n = gridStartIterate(objPos.x, objPos.y, srange, gridListChoose)
            For gi = 0 To n - 1
                psCurr = gridListChoose(gi)
                ' Check that it is a valid target
                c = False
                If SIMPLE_OBJECT_type(psCurr) <> OBJ_FEATURE And SIMPLE_OBJECT_died(psCurr) = 0 Then
                    If Not aiCheckAlliances(SIMPLE_OBJECT_player(psCurr), objPlayer) Then
                        If validTarget(psObj, psCurr, weapon_slot) Then
                            If BASE_OBJECT_visible(psCurr, objPlayer) = UBYTE_MAX Then
                                c = aiStructHasRange(psObj, psCurr, weapon_slot)
                            End If
                        End If
                    End If
                End If
                If c Then
                    newTargetValue = targetAttackWeightIfGreaterThan(targetValue - 1, psCurr, psObj, weapon_slot)
                    ' See if in sensor range and visible
                    distSq = objPosDiffSq_obj(psCurr, psObj)
                    If newTargetValue < targetValue Then
                        GoTo continue_
                    ElseIf newTargetValue = targetValue And distSq >= tarDist Then
                        GoTo continue_
                    End If

                    tmpOrigin = ORIGIN_VISUAL
                    psTarget = psCurr
                    tarDist = distSq
                    targetValue = newTargetValue
                End If
continue_:
            Next
        End If

        If psTarget <> 0 Then
            If Not (SIMPLE_OBJECT_died(psTarget) = 0) Then ASSERT_FAILURE "!psTarget->died", "src/ai.cpp:1004", "aiChooseTarget", "Structure found a dead target!"
            If hasTargetOrigin Then
                targetOrigin = tmpOrigin
            End If
            ppsTarget = psTarget
            aiChooseTarget = True
            Exit Function
        End If
    End If

    aiChooseTarget = False
End Function


' See if there is a target in range for Sensor objects
Public Function aiChooseSensorTarget(ByVal psObj As Long, ppsTarget As Long) As Boolean
    Dim sensorRange As Long, radSquared As Long, radarDetector As Boolean, psTarget As Long, xdiff As Long, ydiff As Long
    Dim distSq As Long, psTemp As Long, tarDist As Double, n As Long, gi As Long, psCurr As Long, objPos As Vector3i
    Dim cpos As Vector3i, tpos As Vector3i, objPlayer As Long

    sensorRange = objSensorRange(psObj)
    radSquared = MUL32(sensorRange, sensorRange)
    radarDetector = objRadarDetector(psObj)

    If objActiveRadar(psObj) = -1 And Not radarDetector Then
        ASSERT_FAILURE "false", "src/ai.cpp:1027", "aiChooseSensorTarget", "Only to be used for sensor turrets!"
        aiChooseSensorTarget = False
        Exit Function
    End If

    objPos = SIMPLE_OBJECT_pos(psObj)
    objPlayer = SIMPLE_OBJECT_player(psObj)

    ' See if there is something in range
    If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        psTarget = 0

        If aiBestNearestTarget(psObj, psTarget, 0) >= 0 Then
            ' See if in sensor range
            tpos = SIMPLE_OBJECT_pos(psTarget)
            xdiff = tpos.x - objPos.x
            ydiff = tpos.y - objPos.y
            distSq = xdiff * xdiff + ydiff * ydiff

            ' I do believe this will never happen, check for yourself :-)
            debug_ 1044, LOG_NEVER, "aiChooseSensorTarget", "Sensor droid(%d) found possible target(%d)!!!", SIMPLE_OBJECT_id(psObj), SIMPLE_OBJECT_id(psTarget)

            If ULT(distSq, radSquared) Then
                ppsTarget = psTarget
                aiChooseSensorTarget = True
                Exit Function
            End If
        End If
    Else    ' structure
        psTemp = 0
        tarDist = 4294967295#   ' UINT32_MAX

        n = gridStartIterate(objPos.x, objPos.y, objSensorRange(psObj), gridListSensor)
        For gi = 0 To n - 1
            psCurr = gridListSensor(gi)
            If psCurr = 0 Then
                GoTo continue_
            End If
            ' Don't target features or doomed/dead objects
            If SIMPLE_OBJECT_type(psCurr) <> OBJ_FEATURE And SIMPLE_OBJECT_died(psCurr) = 0 Then
                If Not aiObjectIsProbablyDoomed(psCurr, False) Then
                    If Not aiCheckAlliances(SIMPLE_OBJECT_player(psCurr), objPlayer) And Not aiObjIsWall(psCurr) Then
                        ' See if in sensor range and visible
                        cpos = SIMPLE_OBJECT_pos(psCurr)
                        xdiff = cpos.x - objPos.x
                        ydiff = cpos.y - objPos.y
                        distSq = xdiff * xdiff + ydiff * ydiff

                        If ULT(distSq, radSquared) And BASE_OBJECT_visible(psCurr, objPlayer) = UBYTE_MAX And U32(distSq) < tarDist Then
                            psTemp = psCurr
                            tarDist = U32(distSq)
                        End If
                    End If
                End If
            End If
continue_:
        Next

        If psTemp <> 0 Then
            If Not (SIMPLE_OBJECT_died(psTemp) = 0) Then ASSERT_FAILURE "!psTemp->died", "src/ai.cpp:1088", "aiChooseSensorTarget", "aiChooseSensorTarget gave us a dead target"
            ppsTarget = psTemp
            aiChooseSensorTarget = True
            Exit Function
        End If
    End If

    aiChooseSensorTarget = False
End Function

' Make droid/structure look for a better target
Private Function updateAttackTarget(ByVal psAttacker As Long, ByVal weapon_slot As Long) As Boolean
    Dim psBetterTarget As Long, tmpOrigin As TARGET_ORIGIN, psDroid As Long, psBuilding As Long, c As Boolean

    psBetterTarget = 0
    tmpOrigin = ORIGIN_UNKNOWN

    If aiChooseTarget(psAttacker, psBetterTarget, weapon_slot, True, True, tmpOrigin) Then  'update target
        If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
            psDroid = psAttacker

            c = orderState(psDroid, DORDER_NONE)
            If Not c Then c = orderState(psDroid, DORDER_GUARD)
            If Not c Then c = orderState(psDroid, DORDER_ATTACKTARGET)
            If c And weapon_slot = 0 Then
                actionDroid_obj psAttacker, DACTION_ATTACK, psBetterTarget
            Else    'can't override current order
                setDroidActionTarget psDroid, psBetterTarget, weapon_slot
            End If
        ElseIf SIMPLE_OBJECT_type(psAttacker) = OBJ_STRUCTURE Then
            psBuilding = psAttacker

            setStructureTarget psBuilding, psBetterTarget, weapon_slot, tmpOrigin
        End If

        updateAttackTarget = True
        Exit Function
    End If

    updateAttackTarget = False
End Function

' Do the AI for a droid
Public Sub aiUpdateDroid(ByVal psDroid As Long)
    Dim lookForTarget As Boolean, updateTarget As Boolean, i As Long, psTarget As Long, c As Boolean

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/ai.cpp:1139", "aiUpdateDroid", "Invalid droid pointer"
    If psDroid = 0 Then
        Exit Sub
    ElseIf isDead(psDroid) Then
        Exit Sub
    End If

    If aDroid(psDroid).droidType <> DROID_SENSOR And aDroid(psDroid).numWeaps = 0 Then
        Exit Sub
    End If

    lookForTarget = False
    updateTarget = False

    ' look for a target if doing nothing
    If orderState(psDroid, DORDER_NONE) Or _
       orderState(psDroid, DORDER_GUARD) Or _
       orderState(psDroid, DORDER_HOLD) Then
        lookForTarget = True
    End If
    ' but do not choose another target if doing anything while guarding
    ' exception for sensors, to allow re-targetting when target is doomed
    If orderState(psDroid, DORDER_GUARD) And aDroid(psDroid).action <> DACTION_NONE And aDroid(psDroid).droidType <> DROID_SENSOR Then
        lookForTarget = False
    End If
    ' don't look for a target if sulking
    If aDroid(psDroid).action = DACTION_SULK Then
        lookForTarget = False
    End If

    ' Only try to update target if already have some target
    If aDroid(psDroid).action = DACTION_ATTACK Or _
       aDroid(psDroid).action = DACTION_MOVEFIRE Or _
       aDroid(psDroid).action = DACTION_MOVETOATTACK Or _
       aDroid(psDroid).action = DACTION_ROTATETOATTACK Then
        updateTarget = True
    End If
    If orderState(psDroid, DORDER_OBSERVE) Or orderState(psDroid, DORDER_ATTACKTARGET) Then
        If aDroid(psDroid).order.psObj <> 0 Then
            If SIMPLE_OBJECT_died(aDroid(psDroid).order.psObj) <> 0 Then
                lookForTarget = True
                updateTarget = False
            End If
        End If
    End If

    ' Don't update target if we are sent to attack and reached attack destination (attacking our target)
    If orderState(psDroid, DORDER_ATTACK) And aDroid(psDroid).psActionTarget(0) = aDroid(psDroid).order.psObj Then
        updateTarget = False
    End If

    ' don't look for a target if there are any queued orders
    If aDroid(psDroid).listSize > 0 Then
        lookForTarget = False
        updateTarget = False
    End If

    ' don't allow units to start attacking if they will switch to guarding the commander
    ' except for sensors: they still look for targets themselves, because
    ' they have wider view
    If hasCommander(psDroid) And aDroid(psDroid).droidType <> DROID_SENSOR Then
        lookForTarget = False
        updateTarget = False
    End If

    If bMultiPlayer And DROID_isVtol(psDroid) And isHumanPlayer(aDroid(psDroid).player) Then
        lookForTarget = False
        updateTarget = False
    End If

    ' CB and VTOL CB droids can't autotarget.
    If aDroid(psDroid).droidType = DROID_SENSOR And Not standardSensorDroid(psDroid) Then
        lookForTarget = False
        updateTarget = False
    End If

    ' do not attack if the attack level is wrong
    If secondaryGetState(psDroid, DSO_ATTACK_LEVEL) <> DSS_ALEV_ALWAYS Then
        lookForTarget = False
    End If

    ' For commanders and non-assigned non-commanders: look for a better target once in a while
    c = False
    If Not lookForTarget And updateTarget And aDroid(psDroid).numWeaps > 0 Then
        If Not hasCommander(psDroid) Then
            c = Fix(U32(FromU32(CDbl(aDroid(psDroid).id) + gameTime)) / TARGET_UPD_SKIP_FRAMES) <> _
                Fix(U32(FromU32(CDbl(aDroid(psDroid).id) + gameTime - deltaGameTime)) / TARGET_UPD_SKIP_FRAMES)
        End If
    End If
    If c Then
        For i = 0 To aDroid(psDroid).numWeaps - 1
            updateAttackTarget psDroid, i
        Next
    End If

    ' Null target - see if there is an enemy to attack

    If lookForTarget And Not updateTarget Then
        If aDroid(psDroid).droidType = DROID_SENSOR Then
            If aiChooseSensorTarget(psDroid, psTarget) Then
                If Not orderState(psDroid, DORDER_HOLD) _
                   And secondaryGetState(psDroid, DSO_HALTTYPE) = DSS_HALT_PURSUE Then
                    aDroid(psDroid).order = DroidOrder_obj(DORDER_OBSERVE, psTarget)
                End If
                actionDroid_obj psDroid, DACTION_OBSERVE, psTarget
            End If
        Else
            c = False
            If IS_TIME_TO_CHECK_FOR_NEW_TARGET(psDroid) Then
                c = aiChooseTarget(psDroid, psTarget, 0, True)
            End If
            If c Then
                If Not orderState(psDroid, DORDER_HOLD) _
                   And secondaryGetState(psDroid, DSO_HALTTYPE) = DSS_HALT_PURSUE Then
                    aDroid(psDroid).order = DroidOrder_obj(DORDER_ATTACK, psTarget)
                End If
                actionDroid_obj psDroid, DACTION_ATTACK, psTarget
            End If
        End If
    End If
End Sub

' Check if any of our weapons can hit the target...
Public Function checkAnyWeaponsTarget(ByVal psObject As Long, ByVal psTarget As Long) As Boolean
    Dim psDroid As Long, i As Long
    psDroid = psObject
    For i = 0 To aDroid(psDroid).numWeaps - 1
        If validTarget(psObject, psTarget, i) Then
            checkAnyWeaponsTarget = True
            Exit Function
        End If
    Next
    checkAnyWeaponsTarget = False
End Function

' Set of rules which determine whether the weapon associated with the object can fire on the propulsion type of the target.
Public Function validTarget(ByVal psObject As Long, ByVal psTarget As Long, ByVal weapon_slot As Long) As Boolean
    Dim bTargetInAir As Boolean, bValidTarget As Boolean, surfaceToAir As Long

    bTargetInAir = False
    bValidTarget = False
    surfaceToAir = 0

    If psTarget = 0 Then
        validTarget = False
        Exit Function
    End If

    ' Need to check propulsion type of target
    Select Case SIMPLE_OBJECT_type(psTarget)
    Case OBJ_DROID
        If asPropulsionTypes(asPropulsionStats(DROID_getPropulsionStats(psTarget)).propulsionType).travel = AIR Then
            If aDroid(psTarget).sMove.Status <> MOVEINACTIVE Then
                bTargetInAir = True
            Else
                bTargetInAir = False
            End If
        Else
            bTargetInAir = False
        End If
    Case Else   ' OBJ_STRUCTURE, default
        'lets hope so!
        bTargetInAir = False
    End Select

    'need what can shoot at
    Select Case SIMPLE_OBJECT_type(psObject)
    Case OBJ_DROID
        If aDroid(psObject).droidType = DROID_SENSOR Then
            validTarget = Not bTargetInAir  ' Sensor droids should not target anything in the air.
            Exit Function
        End If

        ' Can't attack without a weapon
        If aDroid(psObject).numWeaps <> 0 And aDroid(psObject).asWeaps(weapon_slot).nStat <> 0 Then
            surfaceToAir = asWeaponStats(DROID_getWeaponStats(psObject, weapon_slot)).surfaceToAir
            If ((surfaceToAir And SHOOT_IN_AIR) <> 0 And bTargetInAir) Or ((surfaceToAir And SHOOT_ON_GROUND) <> 0 And Not bTargetInAir) Then
                validTarget = True
                Exit Function
            End If
        Else
            validTarget = False
            Exit Function
        End If
    Case OBJ_STRUCTURE
        ' Can't attack without a weapon
        If aStruct(psObject).numWeaps <> 0 And aStruct(psObject).asWeaps(weapon_slot).nStat <> 0 Then
            surfaceToAir = asWeaponStats(BASE_OBJECT_getWeaponStats(psObject, weapon_slot)).surfaceToAir
        Else
            surfaceToAir = 0
        End If

        If ((surfaceToAir And SHOOT_IN_AIR) <> 0 And bTargetInAir) Or ((surfaceToAir And SHOOT_ON_GROUND) <> 0 And Not bTargetInAir) Then
            validTarget = True
            Exit Function
        End If
    Case Else
        surfaceToAir = 0
    End Select

    'if target is in the air and you can shoot in the air - OK
    If bTargetInAir And (surfaceToAir And SHOOT_IN_AIR) <> 0 Then
        bValidTarget = True
    End If

    'if target is on the ground and can shoot at it - OK
    If Not bTargetInAir And (surfaceToAir And SHOOT_ON_GROUND) <> 0 Then
        bValidTarget = True
    End If

    validTarget = bValidTarget
End Function
