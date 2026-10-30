Attribute VB_Name = "combat_c"
' src/combat.cpp, combat.h: Combat mechanics routines. combFire's WEAPON
' *psWeap is a ByRef WEAPON (the attacker's asWeaps element); a BASE_OBJECT's
' timeLastHit, lastHitWeapon, periodicalDamage and body go through
' basedef_h's BASE_OBJECT_<field> properties.
Option Explicit

' (VB) DROID_SHIELD_DAMAGE_SPREAD: (16 - rand()%32); DROID_SHIELD_PARTICLES: (6 + rand()%8)
Private Function DROID_SHIELD_DAMAGE_SPREAD() As Long
    DROID_SHIELD_DAMAGE_SPREAD = (16 - rand() Mod 32)
End Function

Private Function DROID_SHIELD_PARTICLES() As Long
    DROID_SHIELD_PARTICLES = (6 + rand() Mod 8)
End Function

' (VB) statsdef.h's getStatsName for a WEAPON_STATS
Private Function weaponStatsName(ByVal psStats As Long) As String
    weaponStatsName = asWeaponStats(psStats).name_
End Function

' Fire a weapon at something
Public Function combFire(psWeap As WEAPON, ByVal psAttacker As Long, ByVal psTarget As Long, ByVal weapon_slot As Long) As Boolean
    Dim psStats As Long, firePause As Long, longRange As Long, shortRange As Long, compIndex As Long
    Dim fireTime As Long, reloadTime As Long, tall As Boolean, deltaPos As Vector3i, attackerPos As Vector3i, targetPos As Vector3i
    Dim targetDir As Long, dirDiff As Long, dist As Long, min_angle As Long, baseHitChance As Long, min_range As Long
    Dim resultHitChance As Long, level As Long, predict As Vector3i, psDroid As Long, flightTime As Long, vXY As Long, vZ As Long
    Dim empTime As Long, bVisibleAnyway As Boolean, isHit As Boolean, targetShape As ObjectShape, deltaPosPredict As Vector3i
    Dim worstShot As Long, num As Long, minOffset As Long, missDist As Long, miss As Vector3i, ang As Long, attackerPlayer As Long

    checkObject psAttacker, "src/combat.cpp:62", "combFire", max_check_object_recursion
    checkObject psTarget, "src/combat.cpp:63", "combFire", max_check_object_recursion
    ' (VB) ASSERT(psWeap != nullptr): psWeap is a reference

    attackerPlayer = SIMPLE_OBJECT_player(psAttacker)

    ' Don't shoot if the weapon_slot of a vtol is empty
    If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
        If DROID_isVtol(psAttacker) Then
            If Not ULT(psWeap.usedAmmo, getNumAttackRuns(psAttacker, weapon_slot)) Then
                objTrace SIMPLE_OBJECT_id(psAttacker), "combFire", "VTOL slot %d is empty", weapon_slot
                combFire = False
                Exit Function
            End If
        End If
    End If

    ' Get the stats for the weapon
    compIndex = psWeap.nStat
    If Not (ULT(compIndex, asWeaponStats_size)) Then ASSERT_FAILURE "compIndex < asWeaponStats.size()", "src/combat.cpp:76", "combFire", "Invalid range referenced for numWeaponStats, %d > %zu", compIndex, asWeaponStats_size: combFire = False: Exit Function
    psStats = compIndex

    ' check valid weapon/prop combination
    If Not validTarget(psAttacker, psTarget, weapon_slot) Then
        combFire = False
        Exit Function
    End If

    fireTime = FromU32(U32(gameTime) - U32(deltaGameTime) + 1)  ' Can fire earliest at the start of the tick.

    ' See if reloadable weapon.
    If asWeaponStats(psStats).upgrade(attackerPlayer).reloadTime <> 0 Then
        reloadTime = FromU32(U32(psWeap.lastFired) + U32(weaponReloadTime(asWeaponStats(psStats), attackerPlayer)))
        If psWeap.ammo = 0 Then  ' Out of ammo?
            If ULT(fireTime, reloadTime) Then fireTime = reloadTime  ' Have to wait for weapon to reload before firing.
            If ULT(gameTime, fireTime) Then
                combFire = False
                Exit Function
            End If
        End If

        If Not ULT(fireTime, reloadTime) Then
            'reset the ammo level
            psWeap.ammo = asWeaponStats(psStats).upgrade(attackerPlayer).numRounds
        End If
    End If

    ' See when the weapon last fired to control it's rate of fire
    firePause = weaponFirePause(asWeaponStats(psStats), attackerPlayer)
    If ULT(firePause, 1) Then firePause = 1  ' Don't shoot infinitely many shots at once.
    If ULT(fireTime, FromU32(U32(psWeap.lastFired) + U32(firePause))) Then fireTime = FromU32(U32(psWeap.lastFired) + U32(firePause))

    If ULT(gameTime, fireTime) Then
        ' Too soon to fire again
        combFire = False
        Exit Function
    End If

    If Not (attackerPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "psAttacker->player < 11", "src/combat.cpp:118", "combFire", "psAttacker->player = %u", attackerPlayer
    If BASE_OBJECT_visible(psTarget, attackerPlayer) <> UBYTE_MAX Then
        ' Can't see it - can't hit it
        objTrace SIMPLE_OBJECT_id(psAttacker), "combFire", "combFire(%u[%s]->%u): Object has no indirect sight of target", SIMPLE_OBJECT_id(psAttacker), weaponStatsName(psStats), SIMPLE_OBJECT_id(psTarget)
        combFire = False
        Exit Function
    End If

    ' Check we can hit the target
    tall = False
    If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
        tall = DROID_isVtol(psAttacker)
    End If
    If Not tall And SIMPLE_OBJECT_type(psAttacker) = OBJ_STRUCTURE Then
        tall = asStructureStats(aStruct(psAttacker).pStructureType).height > 1
    End If
    If proj_Direct(psStats) Then
        If Not lineOfFire(psAttacker, psTarget, weapon_slot, tall) Then
            ' Can't see the target - can't hit it with direct fire
            objTrace SIMPLE_OBJECT_id(psAttacker), "combFire", "combFire(%u[%s]->%u): No direct line of sight to target", _
                     SIMPLE_OBJECT_id(psAttacker), objInfo(psAttacker), SIMPLE_OBJECT_id(psTarget)
            combFire = False
            Exit Function
        End If
    End If

    attackerPos = SIMPLE_OBJECT_pos(psAttacker)
    targetPos = SIMPLE_OBJECT_pos(psTarget)
    deltaPos.x = targetPos.x - attackerPos.x
    deltaPos.y = targetPos.y - attackerPos.y
    deltaPos.z = targetPos.z - attackerPos.z

    ' if the turret doesn't turn, check if the attacker is in alignment with the target
    If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID And asWeaponStats(psStats).rotate = 0 Then
        targetDir = iAtan2(deltaPos.x, deltaPos.y) And &HFFFF&
        dirDiff = Abs(angleDelta(targetDir - SIMPLE_OBJECT_rot(psAttacker).direction))
        If dirDiff > FIXED_TURRET_DIR Then
            combFire = False
            Exit Function
        End If
    End If

    ' Now see if the target is in range  - also check not too near
    dist = iHypot(deltaPos.x, deltaPos.y)
    longRange = proj_GetLongRange(asWeaponStats(psStats), attackerPlayer)
    shortRange = proj_GetShortRange(asWeaponStats(psStats), attackerPlayer)

    min_angle = 0
    ' Calculate angle for indirect shots
    If Not proj_Direct(psStats) And dist > 0 Then
        min_angle = arcOfFire(psAttacker, psTarget, weapon_slot, True)

        ' prevent extremely steep shots
        min_angle = MIN_(min_angle, DEG(PROJ_ULTIMATE_PITCH))

        ' adjust maximum range of unit if forced to shoot very steep
        If min_angle > DEG(PROJ_MAX_PITCH) Then
            'do not allow increase of max range though
            If iSin(2 * min_angle) < iSin(2 * DEG(PROJ_MAX_PITCH)) Then  ' If PROJ_MAX_PITCH == 45, then always iSin(2*min_angle) <= iSin(2*DEG(PROJ_MAX_PITCH)), and the test is redundant.
                longRange = MUL32(longRange, iSin(2 * min_angle)) \ iSin(2 * DEG(PROJ_MAX_PITCH))
            End If
        End If
    End If

    baseHitChance = 0
    min_range = proj_GetMinRange(asWeaponStats(psStats), attackerPlayer)
    If dist <= shortRange And dist >= min_range Then
        ' get weapon chance to hit in the short range
        baseHitChance = weaponShortHit(asWeaponStats(psStats), attackerPlayer)
    ElseIf dist <= longRange And dist >= min_range Then
        ' get weapon chance to hit in the long range
        baseHitChance = weaponLongHit(asWeaponStats(psStats), attackerPlayer)
    Else
        ' Out of range
        objTrace SIMPLE_OBJECT_id(psAttacker), "combFire", "combFire(%u[%s]->%u): Out of range", SIMPLE_OBJECT_id(psAttacker), weaponStatsName(psStats), SIMPLE_OBJECT_id(psTarget)
        combFire = False
        Exit Function
    End If

    ' adapt for height adjusted artillery shots
    If min_angle > DEG(PROJ_MAX_PITCH) Then
        baseHitChance = baseHitChance * iCos(min_angle) \ iCos(DEG(PROJ_MAX_PITCH))
    End If

    ' apply experience accuracy modifiers to the base
    'hit chance, not to the final hit chance
    resultHitChance = baseHitChance

    ' add the attacker's experience
    If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
        level = getDroidEffectiveLevel(psAttacker, True)

        ' increase total accuracy by EXP_ACCURACY_BONUS % for each experience level
        resultHitChance = resultHitChance + EXP_ACCURACY_BONUS * level * baseHitChance \ 100
    End If

    ' subtract the defender's experience
    If SIMPLE_OBJECT_type(psTarget) = OBJ_DROID Then
        level = getDroidEffectiveLevel(psTarget, True)

        ' decrease weapon accuracy by EXP_ACCURACY_BONUS % for each experience level
        resultHitChance = resultHitChance - EXP_ACCURACY_BONUS * level * baseHitChance \ 100
    End If

    If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
        If aDroid(psAttacker).sMove.Status <> MOVEINACTIVE _
           And Not asWeaponStats(psStats).fireOnMove Then
            combFire = False  ' Can't fire while moving
            Exit Function
        End If
    End If

    ' -------!!! From that point we are sure that we are firing !!!-------

    ' note when the weapon fired
    psWeap.lastFired = fireTime

    ' reduce ammo if salvo
    If asWeaponStats(psStats).upgrade(attackerPlayer).reloadTime <> 0 Then
        psWeap.ammo = FromU32(U32(psWeap.ammo) - 1)
    End If

    ' increment the shots counter
    psWeap.shotsFired = FromU32(U32(psWeap.shotsFired) + 1)

    ' predicted X,Y offset per sec
    predict = targetPos

    ' Target prediction
    If isDroid(psTarget) Then
        If aDroid(castDroid(psTarget)).sMove.bumpTime = 0 Then
            psDroid = castDroid(psTarget)

            If proj_Direct(psStats) Or dist <= proj_GetMinRange(asWeaponStats(psStats), attackerPlayer) Then
                flightTime = FromU32(Fix(U32(MUL32(dist, GAME_TICKS_PER_SEC)) / U32(asWeaponStats(psStats).flightSpeed)))
            Else
                ' Unused, we just want the flight time.
                flightTime = projCalcIndirectVelocities(dist, deltaPos.z, asWeaponStats(psStats).flightSpeed, vXY, vZ, min_angle)
            End If

            If aDroid(psTarget).lastHitWeapon = WSC_EMP Then
                empTime = FromU32(CDbl(EMP_DISABLE_TIME) - (U32(gameTime) - U32(aDroid(psTarget).timeLastHit)))
                CLIP empTime, 0, EMP_DISABLE_TIME
                If empTime >= EMP_DISABLE_TIME * 9 \ 10 Then
                    flightTime = 0  ' Just hit.  Assume they'll get hit again
                Else
                    flightTime = MAX_(0, flightTime - empTime)
                End If
            End If

            ang = aDroid(psDroid).sMove.moveDir
            predict.x = predict.x + iSinR(ang, MUL32(aDroid(psDroid).sMove.speed, flightTime) \ GAME_TICKS_PER_SEC)
            predict.y = predict.y + iCosR(ang, MUL32(aDroid(psDroid).sMove.speed, flightTime) \ GAME_TICKS_PER_SEC)
            If Not DROID_isFlying(psDroid) And Not DROID_isFlightBasedTransporter(psDroid) Then
                predict.z = map_Height(predict.x, predict.y)  ' Predict that the object will be on the ground.
            End If
        End If
    End If

    ' Fire off the bullet to the miss location. The miss is only visible if the player owns the target. (Why? - Per)
    ' What bVisible really does is to make the projectile audible even if it misses you. Since the target is NULL, proj_SendProjectile can't check if it was fired at you.
    bVisibleAnyway = SIMPLE_OBJECT_player(psTarget) = selectedPlayer

    ' see if we were lucky to hit the target
    isHit = gameRand(100) <= resultHitChance
    If isHit Then
        ' Kerrrbaaang !!!!! a hit
        objTrace SIMPLE_OBJECT_id(psAttacker), "combFire", "combFire: [%s]->%u: resultHitChance=%d, visibility=%d", weaponStatsName(psStats), SIMPLE_OBJECT_id(psTarget), resultHitChance, BASE_OBJECT_visible(psTarget, attackerPlayer)
        syncDebug "combFire", "hit=(%d,%d,%d)", predict.x, predict.y, predict.z
    Else ' Deal with a missed shot
        ' Get the shape of the target to avoid "missing" inside of it
        targetShape = establishTargetShape(psTarget)

        ' Worst possible shot based on distance and weapon accuracy
        deltaPosPredict.x = attackerPos.x - predict.x
        deltaPosPredict.y = attackerPos.y - predict.y
        deltaPosPredict.z = attackerPos.z - predict.z
        If resultHitChance > 0 Then
            worstShot = (iHypot(deltaPosPredict.x, deltaPosPredict.y) * 100 \ resultHitChance) \ 5
        Else
            worstShot = iHypot(deltaPosPredict.x, deltaPosPredict.y) * 2
        End If

        ' Use a random seed to determine how far the miss will land from the target
        ' That (num/100)^3 allow the misses to fall much more frequently close to the target
        num = gameRand(100) + 1
        minOffset = 2 * ObjectShape_getRadius(targetShape)

        ' (VB) int arithmetic that wraps, as on x86
        missDist = minOffset + MUL32(MUL32(MUL32(worstShot, num), num), num) \ (100& * 100 * 100)

        ' Determine the angle of the miss in the 270 degrees in "front" of the target.
        ' The 90 degrees behind would most probably cause an unwanted hit when the projectile will be drawn through the hitbox.
        ang = gameRand(DEG(270)) - DEG(135) + iAtan2(deltaPosPredict.x, deltaPosPredict.y)
        miss = Vector3i_(iSinR(ang, missDist), iCosR(ang, missDist), 0)
        predict.x = predict.x + miss.x
        predict.y = predict.y + miss.y
        predict.z = predict.z + miss.z

        psTarget = 0  ' Missed the target, so don't expect to hit it.

        objTrace SIMPLE_OBJECT_id(psAttacker), "combFire", "combFire: Missed shot by (%4d,%4d)", miss.x, miss.y
        syncDebug "combFire", "miss=(%d,%d,%d)", predict.x, predict.y, predict.z
    End If

    ' Make sure we don't pass any negative or out of bounds numbers to proj_SendProjectile
    CLIP predict.x, 0, world_coord(mapWidth - 1)
    CLIP predict.y, 0, world_coord(mapHeight - 1)

    proj_SendProjectileAngled psWeap, psAttacker, attackerPlayer, predict, psTarget, bVisibleAnyway, weapon_slot, min_angle, fireTime
    combFire = True
End Function

' checks through the target players list of structures and droids to see
' if any support a counter battery sensor
Public Sub counterBatteryFire(ByVal psAttacker As Long, ByVal psTarget As Long)
    Dim it As Long, psViewer As Long, sensorRange As Long, xDiff As Long, yDiff As Long, viewerPos As Vector3i, targetPos As Vector3i
    Dim skip As Boolean

    ' if a null target is passed in ignore - this will be the case when a 'miss'
    ' projectile is sent - we may have to cater for these at some point
    ' also ignore cases where you attack your own player
    ' Also ignore cases where there are already 1000 missiles heading towards the attacker.
    If psTarget = 0 Then
        Exit Sub
    End If
    If psAttacker <> 0 Then
        If SIMPLE_OBJECT_player(psAttacker) = SIMPLE_OBJECT_player(psTarget) Then
            Exit Sub
        End If
    End If
    If psAttacker <> 0 Then
        If aiObjectIsProbablyDoomed(psAttacker, False) Then
            Exit Sub
        End If
    End If

    checkObject psTarget, "src/combat.cpp:352", "counterBatteryFire", max_check_object_recursion

    it = apsSensorList(0).head
    Do While it <> 0
        psViewer = lnObj(it)
        If aiCheckAlliances(SIMPLE_OBJECT_player(psTarget), SIMPLE_OBJECT_player(psViewer)) Then
            skip = False
            If SIMPLE_OBJECT_type(psViewer) = OBJ_STRUCTURE Then
                If Not structCBSensor(psViewer) Then
                    skip = asSensorStats(asStructureStats(aStruct(psViewer).pStructureType).pSensor).typ <> VTOL_CB_SENSOR
                End If
            End If
            If Not skip And SIMPLE_OBJECT_type(psViewer) = OBJ_DROID Then
                skip = Not cbSensorDroid(psViewer)
            End If
            If skip Then
                GoTo continue_
            End If
            sensorRange = objSensorRange(psViewer)

            ' Check sensor distance from target
            viewerPos = SIMPLE_OBJECT_pos(psViewer)
            targetPos = SIMPLE_OBJECT_pos(psTarget)
            xDiff = viewerPos.x - targetPos.x
            yDiff = viewerPos.y - targetPos.y

            If xDiff * xDiff + yDiff * yDiff < sensorRange * sensorRange Then
                ' Inform viewer of target
                If SIMPLE_OBJECT_type(psViewer) = OBJ_DROID Then
                    orderDroidObj psViewer, DORDER_OBSERVE, psAttacker, ModeImmediate
                ElseIf SIMPLE_OBJECT_type(psViewer) = OBJ_STRUCTURE Then
                    setStructureTarget psViewer, psAttacker, 0, ORIGIN_CB_SENSOR
                End If
            End If
        End If
continue_:
        it = lnNext(it)
    Loop
End Sub

Public Function objArmour(ByVal psObj As Long, ByVal weaponClass As WEAPON_CLASS) As Long
    Dim armour As Long
    armour = 0
    If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        armour = bodyArmour(asBodyStats(DROID_getBodyStats(psObj)), SIMPLE_OBJECT_player(psObj), weaponClass)
    ElseIf SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE And weaponClass = WC_KINETIC And aStruct(psObj).status <> SS_BEING_BUILT Then
        armour = asStructureStats(aStruct(psObj).pStructureType).upgrade(SIMPLE_OBJECT_player(psObj)).armour
    ElseIf SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE And weaponClass = WC_HEAT And aStruct(psObj).status <> SS_BEING_BUILT Then
        armour = asStructureStats(aStruct(psObj).pStructureType).upgrade(SIMPLE_OBJECT_player(psObj)).thermal
    ElseIf SIMPLE_OBJECT_type(psObj) = OBJ_FEATURE And weaponClass = WC_KINETIC Then
        armour = asFeatureStats(aFeature(psObj).psStats).armourValue
    End If
    objArmour = armour
End Function

' Deals damage to an object
' \param psObj object to deal damage to
' \param psProjectile projectile which hit the object (may be nullptr)
' \param damage amount of damage to deal
' \param weaponClass the class of the weapon that deals the damage
' \param weaponSubClass the subclass of the weapon that deals the damage
' \return < 0 when the dealt damage destroys the object, > 0 when the object survives
Public Function objDamage(ByVal psObj As Long, ByVal psProjectile As Long, ByVal damage As Long, ByVal originalhp As Long, ByVal weaponClass As WEAPON_CLASS, ByVal weaponSubClass As WEAPON_SUBCLASS, ByVal isDamagePerSecond As Boolean, ByVal minDamage As Long, ByVal empRadiusHit As Boolean) As Long
    Dim level As Long, armour As Long, lastHit As Long, psDroid As Long, bMultiMessagesBackup As Boolean
    Dim actualDamage As Long, deltaDamageRate As Long, i As Long, dv As Vector3i, psSource As Long

    level = 0
    armour = objArmour(psObj, weaponClass)
    lastHit = BASE_OBJECT_timeLastHit(psObj)

    ' If the previous hit was by an EMP cannon and this one is not:
    ' don't reset the weapon class and hit time
    ' (Giel: I guess we need this to determine when the EMP-"shock" is over)
    If BASE_OBJECT_lastHitWeapon(psObj) <> WSC_EMP Or weaponSubClass = WSC_EMP Then
        BASE_OBJECT_timeLastHit(psObj) = gameTime
        BASE_OBJECT_lastHitWeapon(psObj) = weaponSubClass
    End If

    ' EMP weapon radius should not do actual damage
    If weaponSubClass = WSC_EMP And empRadiusHit Then
        If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
            psDroid = castDroid(psObj)

            If aDroid(psDroid).shieldPoints > 0 Then
                ' EMP weapons kills droid shields completely
                aDroid(psDroid).shieldPoints = 0
            End If
        End If
        objDamage = 0
        Exit Function
    End If

    ' apply game difficulty setting
    damage = modifyForDifficultyLevel(damage, SIMPLE_OBJECT_player(psObj) <> selectedPlayer)

    If SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE Or SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        ' Force sending messages, even if messages were turned off, since a non-synchronised script will execute here. (Aaargh!)
        bMultiMessagesBackup = bMultiMessages
        bMultiMessages = bMultiPlayer

        psSource = 0
        If psProjectile <> 0 Then psSource = aProj(psProjectile).psSource
        triggerEventAttacked psObj, psSource, lastHit

        bMultiMessages = bMultiMessagesBackup
    End If

    If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        psDroid = psObj

        ' Retrieve highest, applicable, experience level
        level = getDroidEffectiveLevel(psDroid, True)
    ElseIf SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE Then
        level = getStructureDamageBaseExperienceLevel()
    End If

    ' Reduce damage taken by EXP_REDUCE_DAMAGE % for each experience level
    actualDamage = FromU32(Fix(U32(MUL32(damage, 100 - EXP_REDUCE_DAMAGE * level)) / 100))

    ' Apply at least the minimum damage amount
    actualDamage = MAX_(actualDamage - armour, actualDamage * minDamage \ 100)

    ' And at least MIN_WEAPON_DAMAGE points
    actualDamage = MAX_(actualDamage, MIN_WEAPON_DAMAGE)

    debug_ 480, LOG_ATTACK, "objDamage", "objDamage(%d): body: %d, armour: %d, basic damage: %d, actual damage: %d", SIMPLE_OBJECT_id(psObj), BASE_OBJECT_body(psObj), armour, damage, actualDamage

    If isDamagePerSecond Then
        deltaDamageRate = actualDamage - BASE_OBJECT_periodicalDamage(psObj)
        If deltaDamageRate <= 0 Then
            objDamage = 0  ' Did this much damage already, this tick, so don't do more.
            Exit Function
        End If
        actualDamage = gameTimeAdjustedAverage(deltaDamageRate)
        BASE_OBJECT_periodicalDamage(psObj) = BASE_OBJECT_periodicalDamage(psObj) + deltaDamageRate
    End If

    objTrace SIMPLE_OBJECT_id(psObj), "objDamage", "objDamage: Penetrated %d", actualDamage
    syncDebug "objDamage", "damage%u dam%u,o%u,wc%d.%d,ar%d,lev%d,aDam%d,isDps%d", SIMPLE_OBJECT_id(psObj), damage, originalhp, weaponClass, weaponSubClass, armour, level, actualDamage, IIf(isDamagePerSecond, 1, 0)

    ' for some odd reason, we have 0 hitpoints.
    If originalhp = 0 Then
        ASSERT_FAILURE "originalhp", "src/combat.cpp:499", "objDamage", "original hitpoints are 0 ?"
        objDamage = -65536  ' it is dead
        Exit Function
    End If

    ' If the shell did sufficient damage to destroy the object, deal with it and return
    If Not ULT(actualDamage, BASE_OBJECT_body(psObj)) Then
        objDamage = FromU32(Fix(-65536# * U32(BASE_OBJECT_body(psObj)) / U32(originalhp)))
        Exit Function
    End If

    ' Drain shields first
    If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        psDroid = castDroid(psObj)

        If aDroid(psDroid).shieldPoints > 0 Then
            If aDroid(psDroid).shieldPoints > actualDamage Then
                aDroid(psDroid).shieldPoints = aDroid(psDroid).shieldPoints - actualDamage
                actualDamage = 0
            Else
                actualDamage = actualDamage - aDroid(psDroid).shieldPoints
                aDroid(psDroid).shieldPoints = 0
            End If

            If aDroid(psDroid).shieldPoints = 0 Then
                ' shields are interrupted, wait for a while until regeneration starts again
                aDroid(psDroid).shieldInterruptRegenTime = aDroid(psDroid).time_
            End If

            If shieldHitShows(psDroid, psProjectile, weaponSubClass) Then
                dv.y = aProj(psProjectile).pos.z

                For i = 0 To DROID_SHIELD_PARTICLES() - 1
                    dv.x = aProj(psProjectile).pos.x + DROID_SHIELD_DAMAGE_SPREAD()
                    dv.z = aProj(psProjectile).pos.y + DROID_SHIELD_DAMAGE_SPREAD()
                    addEffect dv, EFFECT_FIREWORK, FIREWORK_TYPE_STARBURST, False, 0, 0, gameTime - deltaGameTime + 1
                Next

                audio_PlayStaticTrack aProj(psProjectile).pos.x, aProj(psProjectile).pos.y, ID_SOUND_SHIELD_HIT
            End If
        End If
    End If

    ' Subtract the dealt damage from the droid's remaining body points
    BASE_OBJECT_body(psObj) = FromU32(U32(BASE_OBJECT_body(psObj)) - actualDamage)

    syncDebugObject_ "objDamage", psObj, Asc("D")

    objDamage = FromU32(Fix(65536# * actualDamage / U32(originalhp)))
End Function

' (VB) objDamage: psProjectile != nullptr && weaponSubClass != WSC_FLAME && weaponSubClass != WSC_COMMAND &&
' droidGetMaxShieldPoints(psDroid) > 0 && PERCENT(psDroid->shieldPoints, droidGetMaxShieldPoints(psDroid)) > 25 &&
' objPosDiffSq(psDroid->pos, psProjectile->pos) < TILE_WIDTH * TILE_WIDTH
Private Function shieldHitShows(ByVal psDroid As Long, ByVal psProjectile As Long, ByVal weaponSubClass As WEAPON_SUBCLASS) As Boolean
    If psProjectile = 0 Then Exit Function
    If weaponSubClass = WSC_FLAME Or weaponSubClass = WSC_COMMAND Then Exit Function
    If droidGetMaxShieldPoints(psDroid) <= 0 Then Exit Function
    If Not (FromU32(Fix(U32(MUL32(aDroid(psDroid).shieldPoints, 100)) / U32(droidGetMaxShieldPoints(psDroid)))) > 25) Then Exit Function
    shieldHitShows = objPosDiffSq(aDroid(psDroid).pos, aProj(psProjectile).pos) < TILE_WIDTH * TILE_WIDTH
End Function

' Guesses how damage a shot might do.
' \param psObj object that might be hit
' \param damage amount of damage to deal
' \param weaponClass the class of the weapon that deals the damage
' \param weaponSubClass the subclass of the weapon that deals the damage
' \return guess at amount of damage
' psStats: a WEAPON_STATS * (its index)
Public Function objGuessFutureDamage(ByVal psStats As Long, ByVal player As Long, ByVal psTarget As Long) As Long
    Dim damage As Long, actualDamage As Long, armour As Long, level As Long, psDroid As Long, minDamagePart As Long

    level = 0

    If psTarget = 0 Then
        objGuessFutureDamage = 0  ' Hard to destroy the ground. The armour on the mud is very strong and blocks all damage.
        Exit Function
    End If

    damage = calcDamage(weaponDamage(asWeaponStats(psStats), player), asWeaponStats(psStats).weaponEffect, psTarget)

    ' apply game difficulty setting
    damage = modifyForDifficultyLevel(damage, SIMPLE_OBJECT_player(psTarget) <> selectedPlayer)
    armour = objArmour(psTarget, asWeaponStats(psStats).weaponClass)

    If SIMPLE_OBJECT_type(psTarget) = OBJ_DROID Then
        psDroid = psTarget

        ' Retrieve highest, applicable, experience level
        level = getDroidEffectiveLevel(psDroid, True)
    ElseIf SIMPLE_OBJECT_type(psTarget) = OBJ_STRUCTURE Then
        level = getStructureDamageBaseExperienceLevel()
    End If
    'debug(LOG_ATTACK, "objGuessFutureDamage(%d): body %d armour %d damage: %d", psObj->id, psObj->body, armour, damage);

    ' Reduce damage taken by EXP_REDUCE_DAMAGE % for each experience level
    actualDamage = FromU32(Fix(U32(MUL32(damage, 100 - EXP_REDUCE_DAMAGE * level)) / 100))

    ' You always do at least a third of the experience modified damage
    ' (VB) minimumDamage is unsigned: MAX compares (and picks) as unsigned, so a negative actualDamage - armour wins
    minDamagePart = FromU32(Fix(U32(MUL32(actualDamage, asWeaponStats(psStats).upgrade(player).minimumDamage)) / 100))
    If ULT(minDamagePart, actualDamage - armour) Then
        actualDamage = actualDamage - armour
    Else
        actualDamage = minDamagePart
    End If

    ' And at least MIN_WEAPON_DAMAGE points
    actualDamage = MAX_(actualDamage, MIN_WEAPON_DAMAGE)

    'objTrace(psObj->id, "objGuessFutureDamage: Would penetrate %d", actualDamage);

    objGuessFutureDamage = actualDamage
End Function
