Attribute VB_Name = "projectile_c"
' src/projectile.cpp, projectile.h: Projectile functions. Ported whole but
' objectShimmy and justBeenHitByEW (the display's shimmer of a unit hit by
' electronic warfare: a glm::mat4 for the renderer, no game state).
' A WEAPON_STATS * is its asWeaponStats index (-1: nullptr); a const
' WEAPON_STATS & is the UDT ByRef. A PROJECTILE * is a handle into aProj
' (see projectiledef_h); globalProjectileStorage is
' paged_entity_container_h's ProjectileContainer, psProjectileList the
' array psProjectileList() plus psProjectileList_size, the iterator
' psProjectileNext an index into it. PROJECTILE's constructor and
' destructor are PROJECTILE_construct and PROJECTILE_destruct, its update()
' is PROJECTILE_update (the spawned projectile's handle, 0 = none).
' (VB) proj_SendProjectileAngledInternal takes its slot in the pool where
' C++ constructs the projectile on the stack and moves it into the storage
' later; nothing in between creates another projectile. A caller must not
' pass an aProj element ByRef to a function that sends a projectile (the
' pool may grow: VB's array would be locked), so the penetration passes a
' copy of the projectile's dst. setProjectileSource (projectile.h's inline
' function) is here; basedef.h's setSpacetime, vector.h's
' Vector3i_InSphere and the BASE_OBJECT fields periodicalDamageStart and
' periodicalDamage are this module's private helpers. ObjectShape's
' constructors are ObjectShape_radius (int radius) and ObjectShape_size
' (Vector2i widthBreadth). struct DAMAGE is DAMAGE_ (its field and the
' locals named damage keep the name).
Option Explicit

Public Const PROJ_MAX_PITCH As Long = 45
Public Const PROJ_ULTIMATE_PITCH As Long = 80

Public Const BURN_TIME As Long = 10000      ' How long an object burns for after leaving a fire.
Public Const BURN_DAMAGE As Long = 15       ' How much damage per second an object takes when it is burning.
Public Const BURN_MIN_DAMAGE As Long = 30   ' Least percentage of damage an object takes when burning.
Public Const ACC_GRAVITY As Long = 1000     ' Downward force against projectiles.

' How long to display a single electronic warfare shimmmer.
Public Const ELEC_DAMAGE_DURATION As Long = GAME_TICKS_PER_SEC \ 5

Public Type ObjectShape
    isRectangular As Boolean            ' True if rectangular, false if circular.
    size As Vector2i                    ' x == y if circular.
End Type

Private Const VTOL_HITBOX_MODIFICATOR As Long = 100

Private Const HOMINGINDIRECT_HEIGHT_MIN As Long = 200
Private Const HOMINGINDIRECT_HEIGHT_MAX As Long = 450

Private experienceGain(0 To MAX_PLAYERS - 1) As Long

Private Type INTERVAL
    begin_ As Long                      ' Time 1 = 0, time 2 = 1024. Or begin >= end if empty.
    end_ As Long
End Type

Private Type DAMAGE_
    psProjectile As Long               ' PROJECTILE *
    psDest As Long                      ' BASE_OBJECT *
    damage As Long                      ' unsigned
    weaponClass As WEAPON_CLASS
    weaponSubClass As WEAPON_SUBCLASS
    impactTime As Long                  ' unsigned
    isDamagePerSecond As Boolean
    minDamage As Long
    empRadiusHit As Boolean
End Type

' Watermelon:they are from droid.c
' The range for neighbouring objects
Private Const PROJ_NEIGHBOUR_RANGE As Long = TILE_UNITS * 4
' used to create a specific ID for projectile objects to facilitate tracking them.
Private Const ProjectileTrackerID As Long = &HDEAD0000
Private projectileTrackerIDIncrement As Long        ' uint32_t

' The list of projectiles in play.
' This intermediate container is needed to ensure that projectiles are always
' enumerated in a stable and predictable order, because `globalProjectileStorage`
' may insert new elements in place of old ones, which were previously destroyed,
' thus, modifying the order of iteration.
Private psProjectileList() As Long      ' std::vector<PROJECTILE*>
Private psProjectileList_size As Long

' The next projectile to give out in the proj_First / proj_Next methods.
' References `psProjectileList` container.
Private psProjectileNext As Long        ' (VB) an index into psProjectileList (psProjectileList_size: end())

Private Const BULLET_FLIGHT_HEIGHT As Long = 16

Private gridListInFlight() As Long      ' proj_InFlightFunc's static GridList
Private gridListSweep() As Long         ' proj_radiusSweep's static GridList
Private gridListPeriodical() As Long    ' proj_checkPeriodicalDamage's static GridList
Private spawnedProjectiles() As Long    ' proj_UpdateAll's static std::vector<PROJECTILE*>

' ---- PROJECTILE's constructor and destructor

' PROJECTILE(uint32_t id, unsigned player) : SIMPLE_OBJECT(OBJ_PROJECTILE, id, player); h: a zeroed slot
Public Sub PROJECTILE_construct(ByVal h As Long, ByVal id As Long, ByVal player As Long)
    ' SIMPLE_OBJECT
    aProj(h).typ = OBJ_PROJECTILE
    aProj(h).id = id
    aProj(h).pos = Vector3i_(0, 0, 0)
    aProj(h).rot = Rotation_(0, 0, 0)
    aProj(h).player = player And &HFF&
    aProj(h).born = gameTime
    aProj(h).died = 0
    aProj(h).time_ = 0
    ' PROJECTILE's member initialisers
    aProj(h).src = Vector3i_(0, 0, 0)
    aProj(h).dst = Vector3i_(0, 0, 0)
    aProj(h).psDamaged_size = 0
End Sub

Public Sub PROJECTILE_destruct(ByVal this As Long)
    ' (the member psDamaged)
    Erase aProj(this).psDamaged
    aProj(this).psDamaged_size = 0
    ' ~SIMPLE_OBJECT
    aProj(this).typ = aProj(this).typ + 1000000000      ' Hopefully this will trigger an assert              if someone uses the freed object.
    aProj(this).player = (aProj(this).player + 100) And &HFF&  ' Hopefully this will trigger an assert and/or crash if someone uses the freed object.
End Sub

' (VB) psDamaged.push_back(psObj)
Private Sub psDamaged_push(ByVal psProj As Long, ByVal psObj As Long)
    Dim n As Long
    n = aProj(psProj).psDamaged_size
    If n = 0 Then
        ReDim aProj(psProj).psDamaged(0 To 3)
    ElseIf n > UBound(aProj(psProj).psDamaged) Then
        ReDim Preserve aProj(psProj).psDamaged(0 To 2 * n - 1)
    End If
    aProj(psProj).psDamaged(n) = psObj
    aProj(psProj).psDamaged_size = n + 1
End Sub

' (VB) std::find(psDamaged.begin(), psDamaged.end(), psObj) != psDamaged.end()
Private Function psDamaged_contains(ByVal psProj As Long, ByVal psObj As Long) As Boolean
    Dim i As Long
    For i = 0 To aProj(psProj).psDamaged_size - 1
        If aProj(psProj).psDamaged(i) = psObj Then
            psDamaged_contains = True
            Exit Function
        End If
    Next
End Function

' (VB) basedef.h's setSpacetime(SIMPLE_OBJECT *psObj, Spacetime const &st), for a projectile
Private Sub setSpacetime(ByVal psProj As Long, st As Spacetime)
    aProj(psProj).pos = st.pos
    aProj(psProj).rot = st.rot
    aProj(psProj).time_ = st.time_
End Sub

' (VB) vector.h: Vector3i_InSphere(Vector3i v, Vector3i c, unsigned r): (unsigned int)dot(delta, delta) < r * r
Private Function Vector3i_InSphere(v As Vector3i, c As Vector3i, ByVal r As Long) As Boolean
    Dim dx As Double, dy As Double, dz As Double
    dx = CDbl(v.x) - c.x
    dy = CDbl(v.y) - c.y
    dz = CDbl(v.z) - c.z
    Vector3i_InSphere = ULT(FromU32(dx * dx + dy * dy + dz * dz), MUL32(r, r))
End Function

' (VB) a BASE_OBJECT's periodicalDamageStart (basedef_h has periodicalDamage)
Private Function objPeriodicalDamageStart(ByVal psObj As Long) As Long
    Select Case objHandleType(psObj)
    Case OBJ_DROID: objPeriodicalDamageStart = aDroid(psObj).periodicalDamageStart
    Case OBJ_STRUCTURE: objPeriodicalDamageStart = aStruct(psObj).periodicalDamageStart
    Case OBJ_FEATURE: objPeriodicalDamageStart = aFeature(psObj).periodicalDamageStart
    End Select
End Function

Private Sub setObjPeriodicalDamageStart(ByVal psObj As Long, ByVal v As Long)
    Select Case objHandleType(psObj)
    Case OBJ_DROID: aDroid(psObj).periodicalDamageStart = v
    Case OBJ_STRUCTURE: aStruct(psObj).periodicalDamageStart = v
    Case OBJ_FEATURE: aFeature(psObj).periodicalDamageStart = v
    End Select
End Sub

' ObjectShape(int radius)
Public Function ObjectShape_radius(ByVal radius As Long) As ObjectShape
    ObjectShape_radius.isRectangular = False
    ObjectShape_radius.size = Vector2i_(radius, radius)
End Function

' ObjectShape(Vector2i widthBreadth)
Public Function ObjectShape_size(widthBreadth As Vector2i) As ObjectShape
    ObjectShape_size.isRectangular = True
    ObjectShape_size.size = widthBreadth
End Function

' ObjectShape::radius()
Public Function ObjectShape_getRadius(shape As ObjectShape) As Long
    ObjectShape_getRadius = shape.size.x
End Function

' projectile.h
Public Sub setProjectileSource(ByVal psProj As Long, ByVal psObj As Long)
    Dim psPrevProj As Long
    ' use the source of the source of psProj if psAttacker is a projectile
    aProj(psProj).psSource = 0
    If psObj = 0 Then
    ElseIf isProjectile(psObj) Then
        psPrevProj = castProjectile(psObj)

        If aProj(psPrevProj).psSource <> 0 Then
            If SIMPLE_OBJECT_died(aProj(psPrevProj).psSource) = 0 Then
                aProj(psProj).psSource = aProj(psPrevProj).psSource
            End If
        End If
    Else
        aProj(psProj).psSource = castBaseObject(psObj)
    End If
End Sub

Private Sub setProjectileDestination(ByVal psProj As Long, ByVal psObj As Long)
    Dim bDirect As Boolean
    bDirect = proj_Direct(aProj(psProj).psWStats)
    aiObjectAddExpectedDamage aProj(psProj).psDest, FromU32(-U32(aProj(psProj).expectedDamageCaused)), bDirect  ' The old target shouldn't be expecting any more damage from this projectile.
    aProj(psProj).psDest = psObj
    aiObjectAddExpectedDamage aProj(psProj).psDest, aProj(psProj).expectedDamageCaused, bDirect  ' Let the new target know to say its prayers.
End Sub

Public Function gfxVisible(ByVal psObj As Long) As Boolean
    Dim psSource As Long, psDest As Long, c As Boolean

    ' Already know it is visible
    If aProj(psObj).bVisible <> 0 Then
        gfxVisible = True
        Exit Function
    End If

    ' You fired it
    If aProj(psObj).player = selectedPlayer Then
        gfxVisible = True
        Exit Function
    End If

    psSource = aProj(psObj).psSource
    psDest = aProj(psObj).psDest

    ' Someone elses structure firing at something you can't see
    If psSource <> 0 Then
        If SIMPLE_OBJECT_died(psSource) = 0 _
           And SIMPLE_OBJECT_type(psSource) = OBJ_STRUCTURE _
           And SIMPLE_OBJECT_player(psSource) <> selectedPlayer Then
            c = psDest = 0
            If Not c Then c = SIMPLE_OBJECT_died(psDest) <> 0
            If Not c Then c = BASE_OBJECT_visibleForLocalDisplay(psDest) = 0
            If c Then
                gfxVisible = False
                Exit Function
            End If
        End If
    End If

    ' Something you cannot see firing at a structure that isn't yours
    If psDest <> 0 Then
        If SIMPLE_OBJECT_died(psDest) = 0 _
           And SIMPLE_OBJECT_type(psDest) = OBJ_STRUCTURE _
           And SIMPLE_OBJECT_player(psDest) <> selectedPlayer Then
            c = psSource = 0
            If Not c Then c = BASE_OBJECT_visibleForLocalDisplay(psSource) = 0
            If c Then
                gfxVisible = False
                Exit Function
            End If
        End If
    End If

    ' You can see the source
    If psSource <> 0 Then
        If SIMPLE_OBJECT_died(psSource) = 0 Then
            If BASE_OBJECT_visibleForLocalDisplay(psSource) <> 0 Then
                gfxVisible = True
                Exit Function
            End If
        End If
    End If

    ' You can see the destination
    If psDest <> 0 Then
        If SIMPLE_OBJECT_died(psDest) = 0 Then
            If BASE_OBJECT_visibleForLocalDisplay(psDest) <> 0 Then
                gfxVisible = True
                Exit Function
            End If
        End If
    End If

    gfxVisible = False
End Function

Public Function proj_InitSystem() As Boolean
    Dim x As Long
    psProjectileList_size = 0
    psProjectileNext = psProjectileList_size
    For x = 0 To MAX_PLAYERS - 1
        experienceGain(x) = 100
    Next
    projectileTrackerIDIncrement = 0
    proj_InitSystem = True
End Function

' Add allocated projectile `p` to the list of active projectiles (`psProjectileList`)
Public Sub proj_AddActiveProjectile(ByVal p As Long)
    If psProjectileList_size = 0 Then
        ReDim psProjectileList(0 To 63)
    ElseIf psProjectileList_size > UBound(psProjectileList) Then
        ReDim Preserve psProjectileList(0 To 2 * psProjectileList_size - 1)
    End If
    psProjectileList(psProjectileList_size) = p
    psProjectileList_size = psProjectileList_size + 1
End Sub

' Clean out all projectiles from the system, and properly decrement
' all reference counts.
Public Sub proj_FreeAllProjectiles()
    Dim i As Long
    For i = 0 To psProjectileList_size - 1
        ' Make sure to get rid of some final references in the sound code to this object first
        audio_RemoveObj psProjectileList(i)
    Next

    psProjectileList_size = 0
    psProjectileNext = psProjectileList_size

    ProjectileContainer_clear
End Sub

Public Function proj_Shutdown() As Boolean
    proj_FreeAllProjectiles

    proj_Shutdown = True
End Function

' Reset the first/next methods, and give out the first projectile in the list.
Public Function proj_GetFirst() As Long
    psProjectileNext = 0
    If psProjectileNext <> psProjectileList_size Then proj_GetFirst = psProjectileList(psProjectileNext) Else proj_GetFirst = 0
End Function

' Get the next projectile
Public Function proj_GetNext() As Long
    psProjectileNext = psProjectileNext + 1
    If psProjectileNext <> psProjectileList_size Then proj_GetNext = psProjectileList(psProjectileNext) Else proj_GetNext = 0
End Function

' Relates the quality of the attacker to the quality of the victim.
' The value returned satisfies the following inequality: 0.5 <= ret/65536 <= 2.0
Private Function qualityFactor(ByVal psAttacker As Long, ByVal psVictim As Long) As Long
    Dim attackerPower As Double, attackerPoints As Double, powerRatio As Double, pointsRatio As Double
    attackerPower = U32(calcDroidPower(psAttacker))
    attackerPoints = U32(calcDroidPoints(psAttacker))
    If attackerPower = 0 Or attackerPoints = 0 Then
        qualityFactor = 65536 \ 2
        Exit Function
    End If
    powerRatio = U32(FromU32(Int(65536# * U32(calcDroidPower(psVictim)) / attackerPower)))
    pointsRatio = U32(FromU32(Int(65536# * U32(calcDroidPoints(psVictim)) / attackerPoints)))

    If powerRatio < 65536 \ 2 Then powerRatio = 65536 \ 2
    If powerRatio > 65536 * 2 Then powerRatio = 65536 * 2
    If pointsRatio < 65536 \ 2 Then pointsRatio = 65536 \ 2
    If pointsRatio > 65536 * 2 Then pointsRatio = 65536 * 2

    qualityFactor = Int((powerRatio + pointsRatio) / 2)
End Function

Public Sub setExpGain(ByVal player As Long, ByVal gain As Long)
    experienceGain(player) = gain
End Sub

Public Function getExpGain(ByVal player As Long) As Long
    getExpGain = experienceGain(player)
End Function

Public Function getDesignatorAttackingObject(ByVal player As Long, ByVal target As Long) As Long
    Dim psCommander As Long
    psCommander = cmdDroidGetDesignator(player)

    getDesignatorAttackingObject = 0
    If psCommander <> 0 Then
        If aDroid(psCommander).action = DACTION_ATTACK And aDroid(psCommander).psActionTarget(0) = target Then
            getDesignatorAttackingObject = psCommander
        End If
    End If
End Function

' update the source experience after a target is damaged/destroyed
' experienceInc: uint32_t
Private Sub proj_UpdateExperience(ByVal psObj As Long, ByVal experienceInc As Long)
    Dim psDroid As Long, psSensor As Long, c As Boolean

    checkProjectile psObj, "src/projectile.cpp:321", "proj_UpdateExperience", max_check_object_recursion

    If SIMPLE_OBJECT_type(aProj(psObj).psSource) = OBJ_DROID Then      ' update droid kills
        psDroid = aProj(psObj).psSource

        ' If it is 'droid-on-droid' then modify the experience by the Quality factor
        ' Only do this in MP so to not un-balance the campaign
        c = False
        If aProj(psObj).psDest <> 0 Then
            If SIMPLE_OBJECT_type(aProj(psObj).psDest) = OBJ_DROID Then c = bMultiPlayer
        End If
        If c Then
            ' Modify the experience gained by the 'quality factor' of the units
            experienceInc = FromU32(Int(U32(experienceInc) * U32(qualityFactor(psDroid, aProj(psObj).psDest)) / 65536))
        End If

        If Not (ULT(experienceInc, 137625)) Then ASSERT_FAILURE "experienceInc < (int)(2.1 * 65536)", "src/projectile.cpp:337", "proj_UpdateExperience", "Experience increase out of range": Exit Sub

        droidIncreaseExperience psDroid, experienceInc

        cmdDroidUpdateExperience psDroid, experienceInc

        psSensor = orderStateObj(psDroid, DORDER_FIRESUPPORT)
        If psSensor <> 0 Then
            If SIMPLE_OBJECT_type(psSensor) = OBJ_DROID Then
                droidIncreaseExperience psSensor, experienceInc
            End If
        End If
    ElseIf SIMPLE_OBJECT_type(aProj(psObj).psSource) = OBJ_STRUCTURE Then
        If Not (ULT(experienceInc, 137625)) Then ASSERT_FAILURE "experienceInc < (int)(2.1 * 65536)", "src/projectile.cpp:351", "proj_UpdateExperience", "Experience increase out of range": Exit Sub

        psDroid = getDesignatorAttackingObject(SIMPLE_OBJECT_player(aProj(psObj).psSource), aProj(psObj).psDest)

        If psDroid <> 0 Then
            droidIncreaseExperience psDroid, experienceInc
        End If
    End If
End Sub

' syncDebugProjectile(psProj, ch) is syncDebugProjectile_ function, psProj, ch
Public Sub syncDebugProjectile_(ByVal function_ As String, ByVal psProj As Long, ByVal ch As Long)
    Dim list(0 To 10) As Long
    If aProj(psProj).typ <> OBJ_PROJECTILE Then
        ASSERT_FAILURE "false", "src/projectile.cpp:367", "_syncDebugProjectile", "%c Broken psProj->type %u!", ch, aProj(psProj).typ
        syncDebug "_syncDebugProjectile", "Broken psProj->type %u!", aProj(psProj).typ
    End If
    list(0) = ch

    list(1) = aProj(psProj).player
    list(2) = aProj(psProj).pos.x
    list(3) = aProj(psProj).pos.y
    list(4) = aProj(psProj).pos.z
    list(5) = aProj(psProj).rot.direction
    list(6) = aProj(psProj).rot.pitch
    list(7) = aProj(psProj).rot.roll
    list(8) = aProj(psProj).state
    list(9) = aProj(psProj).expectedDamageCaused
    list(10) = aProj(psProj).psDamaged_size
    syncDebugIntList_ function_, "%c projectile = p%d;pos(%d,%d,%d),rot(%d,%d,%d),state%d,expectedDamageCaused%d,numberDamaged%u", list, 11
End Sub

Private Function randomVariation(ByVal val As Long) As Long
    ' Up to +-5% random variation.
    randomVariation = FromU32(Fix(CDbl(val) * (95000 + gameRand(10001)) / 100000))
End Function

' (VB) an integer (a Decimal) wrapped to uint64_t, 0 .. 2^64 - 1
Private Function U64(ByVal v As Variant) As Variant
    Dim m As Variant
    m = CDec("18446744073709551616")
    v = CDec(v)
    v = v - Int(v / m) * m
    ' the quotient is rounded to 28 digits: correct by one
    If v < 0 Then v = v + m
    If v >= m Then v = v - m
    U64 = v
End Function

' (VB) an integer (a Decimal) wrapped to int64_t
Private Function I64(ByVal v As Variant) As Variant
    v = U64(v)
    If v >= CDec("9223372036854775808") Then v = v - CDec("18446744073709551616")
    I64 = v
End Function

' (VB) C's integer division of Decimals, truncating toward zero
Private Function DivTrunc(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim na As Variant, nb As Variant, q As Variant
    na = Abs(CDec(a))
    nb = Abs(CDec(b))
    q = Int(na / nb)
    ' the quotient is rounded to 28 digits: correct it
    Do While q * nb > na
        q = q - 1
    Loop
    Do While (q + 1) * nb <= na
        q = q + 1
    Loop
    If (a < 0) <> (b < 0) Then q = -q
    DivTrunc = q
End Function

Public Function projCalcIndirectVelocities(ByVal dx As Long, ByVal dz As Long, ByVal v As Long, vx As Long, vz As Long, ByVal min_angle As Long) As Long
    Dim g As Long, a As Long, b As Variant, c As Variant, t As Long, sq As Long, mytan As Variant, num As Variant

    ' Find values of vx and vz, which solve the equations:
    ' dz = -1/2 g t^2 + vz t
    ' dx = vx t
    ' v^2 = vx^2 + vz^2
    ' Increases v, if needed for there to be a solution. Decreases v, if needed for vz > 0.
    ' Randomly changes v by up to 2.5%, so the shots don't all follow the same path.

    g = ACC_GRAVITY                                                         ' In units/s^2.
    a = FromU32(CDbl(randomVariation(MUL32(v, v))) - CDbl(MUL32(dz, g)))   ' In units^2/s^2.
    b = U64(CDec(MUL32(g, g)) * (CDec(dx) * CDec(dx) + CDec(dz) * CDec(dz)))  ' In units^4/s^4. Casting to uint64_t does sign extend the int32_t.
    c = I64(CDec(a) * CDec(a) - b)                                          ' In units^4/s^4.
    If c < 0 Then
        ' Must increase velocity, target too high. Find the smallest possible a (which corresponds to the smallest possible velocity).

        a = FromU32(CDbl(i64Sqrt(b)) + 1)                                   ' Still in units^2/s^2. Adding +1, since i64Sqrt rounds down.
        c = I64(CDec(a) * CDec(a) - b)                                      ' Still in units^4/s^4. Should be 0, plus possible rounding errors.
    End If

    ' (VB) 2 * (a - i64Sqrt(c)) is an int32_t, iSqrt takes it as a uint32_t
    sq = iSqrt(MUL32(2, FromU32(CDbl(a) - i64Sqrt(U64(c)))))
    t = MAX_(1, MUL32(sq, GAME_TICKS_PER_SEC \ g))                         ' In ticks. Note that a - sqrt(c) >= 0, since c <= a^2. Try changing the - to +, and watch the mini-rockets.
    vx = MUL32(dx, GAME_TICKS_PER_SEC) \ t                                 ' In units/sec.
    vz = FromU32(CDbl(MUL32(dz, GAME_TICKS_PER_SEC) \ t) + MUL32(g, t) \ (2 * GAME_TICKS_PER_SEC))   ' In units/sec.

    ' STATIC_ASSERT(GAME_TICKS_PER_SEC / ACC_GRAVITY * ACC_GRAVITY == GAME_TICKS_PER_SEC)

    If vz < 0 Then
        ' Don't want to shoot downwards, reduce velocity and let gravity take over.
        num = U64(U64(U64(CDec(MUL32(-2, dz))) * GAME_TICKS_PER_SEC) * GAME_TICKS_PER_SEC)
        t = MAX_(1, i64Sqrt(DivTrunc(num, g)))                              ' Still in ticks.
        vx = MUL32(dx, GAME_TICKS_PER_SEC) \ t                             ' Still in units/sec.
        vz = 0                                                              ' Still in units/sec. (Wouldn't really matter if it was pigeons/inch, since it's 0 anyway.)
    End If

    ' CorvusCorax: Check against min_angle
    If iAtan2(vz, vx) < min_angle Then
        ' set pitch to pass terrain
        ' tan(min_angle)=mytan/65536
        mytan = DivTrunc(CDec(iSin(min_angle)) * 65536, iCos(min_angle))
        num = I64(I64(I64(2 * I64(CDec(dx) * mytan - CDec(MUL32(dz, 65536)))) * GAME_TICKS_PER_SEC) * GAME_TICKS_PER_SEC)
        t = MAX_(1, i64Sqrt(U64(DivTrunc(num, CDec(MUL32(g, 65536))))))   ' Still in ticks.
        vx = MUL32(dx, GAME_TICKS_PER_SEC) \ t
        ' mytan=65536*vz/vx
        vz = FromU32Dec(DivTrunc(I64(mytan * vx), 65536))
    End If

    projCalcIndirectVelocities = t
End Function

' (VB) psAttacker: a SIMPLE_OBJECT * (0, a droid, a structure or a projectile); target is not an aProj element (the pool may grow)
Private Function proj_SendProjectileAngledInternal(psWeap As WEAPON, ByVal psAttacker As Long, ByVal player As Long, target As Vector3i, ByVal psTarget As Long, ByVal bVisible As Boolean, ByVal weapon_slot As Long, ByVal min_angle As Long, ByVal fireTime As Long) As Long
    Dim psStats As Long, h As Long, psOldProjectile As Long, i As Long, maxHeight As Long, minHeight As Long
    Dim deltaPos As Vector3i, dist As Long, muzzle As Vector3i, targetPos As Vector3i, c As Boolean

    psStats = psWeap.nStat

    If Not (ULT(psWeap.nStat, asWeaponStats_size)) Then ASSERT_FAILURE "psWeap->nStat < asWeaponStats.size()", "src/projectile.cpp:444", "proj_SendProjectileAngledInternal", "Invalid range referenced for numWeaponStats, %d > %zu", psWeap.nStat, asWeaponStats_size: proj_SendProjectileAngledInternal = 0: Exit Function
    If Not (psStats <> -1) Then ASSERT_FAILURE "psStats != nullptr", "src/projectile.cpp:445", "proj_SendProjectileAngledInternal", "Invalid weapon stats": proj_SendProjectileAngledInternal = 0: Exit Function
    c = psTarget = 0
    If Not c Then c = SIMPLE_OBJECT_died(psTarget) = 0
    If Not c Then ASSERT_FAILURE "psTarget == nullptr || !psTarget->died", "src/projectile.cpp:446", "proj_SendProjectileAngledInternal", "Aiming at dead target!": proj_SendProjectileAngledInternal = 0: Exit Function

    projectileTrackerIDIncrement = FromU32(U32(projectileTrackerIDIncrement) + 1)
    h = ProjectileContainer_alloc()
    PROJECTILE_construct h, FromU32(U32(ProjectileTrackerID) + U32(projectileTrackerIDIncrement)), player

    ' get muzzle offset
    If psAttacker = 0 Then
        ' if there isn't an attacker just start at the target position
        ' NB this is for the script function to fire the las sats
        aProj(h).src = target
    ElseIf SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID And weapon_slot >= 0 Then
        muzzle = aProj(h).src
        calcDroidMuzzleLocation psAttacker, muzzle, weapon_slot
        aProj(h).src = muzzle
        'update attack runs for VTOL droid's each time a shot is fired
        updateVtolAttackRun psAttacker, weapon_slot
    ElseIf SIMPLE_OBJECT_type(psAttacker) = OBJ_STRUCTURE And weapon_slot >= 0 Then
        muzzle = aProj(h).src
        calcStructureMuzzleLocation psAttacker, muzzle, weapon_slot
        aProj(h).src = muzzle
    Else ' incase anything wants a projectile
        aProj(h).src = SIMPLE_OBJECT_pos(psAttacker)
    End If

    ' Initialise the structure
    aProj(h).psWStats = psStats

    aProj(h).pos = aProj(h).src
    aProj(h).dst = target

    aProj(h).bVisible = 0

    ' Must set ->psDest and ->expectedDamageCaused before first call to setProjectileDestination().
    aProj(h).psDest = 0
    aProj(h).expectedDamageCaused = objGuessFutureDamage(psStats, player, psTarget)
    setProjectileDestination h, psTarget  ' Updates expected damage of proj.psDest, using proj.expectedDamageCaused.

    ' When we have been created by penetration (spawned from another projectile),
    ' we shall live no longer than the original projectile may have lived
    c = False
    If psAttacker <> 0 Then c = SIMPLE_OBJECT_type(psAttacker) = OBJ_PROJECTILE
    If c Then
        psOldProjectile = psAttacker
        aProj(h).born = aProj(psOldProjectile).born
        aProj(h).src = aProj(psOldProjectile).src

        aProj(h).prevSpacetime.time_ = aProj(psOldProjectile).time_  ' Have partially ticked already.
        aProj(h).time_ = gameTime
        If aProj(h).prevSpacetime.time_ = aProj(h).time_ Then aProj(h).prevSpacetime.time_ = FromU32(U32(aProj(h).prevSpacetime.time_) - 1)  ' Times should not be equal, for interpolation.

        setProjectileSource h, aProj(psOldProjectile).psSource
        aProj(h).psDamaged_size = 0
        For i = 0 To aProj(psOldProjectile).psDamaged_size - 1
            psDamaged_push h, aProj(psOldProjectile).psDamaged(i)
        Next

        ' TODO Should finish the tick, when penetrating.
    Else
        aProj(h).born = fireTime  ' Born at the start of the tick.

        aProj(h).prevSpacetime.time_ = fireTime
        aProj(h).time_ = aProj(h).prevSpacetime.time_

        setProjectileSource h, psAttacker
    End If

    If psTarget <> 0 Then
        maxHeight = establishTargetHeight(psTarget)
        minHeight = MIN_(MAX_(maxHeight + 2 * LINE_OF_FIRE_MINIMUM - areaOfFire(psAttacker, psTarget, weapon_slot, True), 0), maxHeight)
        scoreUpdateVar WD_SHOTS_ON_TARGET

        targetPos = SIMPLE_OBJECT_pos(psTarget)
        aProj(h).dst.z = targetPos.z + minHeight + gameRand(MAX_(maxHeight - minHeight, 1))
        ' store visible part (LOCK ON this part for homing :)
        aProj(h).partVisible = maxHeight - minHeight
    Else
        If psAttacker = 0 Then
            aProj(h).dst.z = target.z - LINE_OF_FIRE_MINIMUM
        Else
            aProj(h).dst.z = target.z + LINE_OF_FIRE_MINIMUM
        End If
        scoreUpdateVar WD_SHOTS_OFF_TARGET
    End If

    deltaPos.x = aProj(h).dst.x - aProj(h).src.x
    deltaPos.y = aProj(h).dst.y - aProj(h).src.y
    deltaPos.z = aProj(h).dst.z - aProj(h).src.z

    ' roll never set
    aProj(h).rot.roll = 0

    aProj(h).rot.direction = iAtan2(deltaPos.x, deltaPos.y)


    ' Get target distance, horizontal distance only.
    dist = iHypot(deltaPos.x, deltaPos.y)

    If proj_Direct(psStats) Then
        aProj(h).rot.pitch = iAtan2(deltaPos.z, dist)
    Else
        ' indirect
        projCalcIndirectVelocities dist, deltaPos.z, asWeaponStats(psStats).flightSpeed, aProj(h).vXY, aProj(h).vZ, min_angle
        aProj(h).rot.pitch = iAtan2(aProj(h).vZ, aProj(h).vXY)
    End If
    aProj(h).state = PROJ_INFLIGHT

    ' If droid or structure, set muzzle pitch.
    ' Don't allow pitching the muzzle above outside the weapon's limits.
    If psAttacker <> 0 And weapon_slot >= 0 Then
        If SIMPLE_OBJECT_type(psAttacker) = OBJ_DROID Then
            aDroid(psAttacker).asWeaps(weapon_slot).rot.pitch = clip_(angleDelta(aProj(h).rot.pitch), DEG(asWeaponStats(psStats).minElevation), DEG(asWeaponStats(psStats).maxElevation)) And &HFFFF&
        ElseIf SIMPLE_OBJECT_type(psAttacker) = OBJ_STRUCTURE Then
            aStruct(psAttacker).asWeaps(weapon_slot).rot.pitch = clip_(angleDelta(aProj(h).rot.pitch), DEG(asWeaponStats(psStats).minElevation), DEG(asWeaponStats(psStats).maxElevation)) And &HFFFF&
        End If
    End If

    ' put the projectile object in the global list, obtain the stable address for it.
    ' (VB) h: taken at the construction

    ' play firing audio
    ' only play if either object is visible, i know it's a bit of a hack, but it avoids the problem
    ' of having to calculate real visibility values for each projectile.
    c = bVisible
    If Not c Then c = gfxVisible(h)
    If c Then
        ' note that the projectile is visible
        aProj(h).bVisible = 1

        If asWeaponStats(psStats).iAudioFireID <> NO_SOUND Then

            If aProj(h).psSource <> 0 Then
                ' firing sound emitted from source
                audio_PlayObjDynamicTrack aProj(h).psSource, asWeaponStats(psStats).iAudioFireID, 0
                ' GJ HACK: move howitzer sound with shell
                If asWeaponStats(psStats).weaponSubClass = WSC_HOWITZERS Then
                    audio_PlayObjDynamicTrack h, ID_SOUND_HOWITZ_FLIGHT, 0
                End If
            'don't play the sound for a LasSat in multiPlayer
            ElseIf Not (bMultiPlayer And asWeaponStats(psStats).weaponSubClass = WSC_LAS_SAT) Then
                audio_PlayObjStaticTrack h, asWeaponStats(psStats).iAudioFireID
            End If
        End If
    End If

    If psAttacker <> 0 And Not proj_Direct(psStats) Then
        'check for Counter Battery Sensor in range of target
        counterBatteryFire castBaseObject(psAttacker), psTarget
    End If

    syncDebugProjectile_ "proj_SendProjectileAngledInternal", h, Asc("*")

    checkProjectile h, "src/projectile.cpp:614", "proj_SendProjectileAngledInternal", max_check_object_recursion

    proj_SendProjectileAngledInternal = h
End Function

Private Function proj_SendProjectileInternal(psWeap As WEAPON, ByVal psAttacker As Long, ByVal player As Long, target As Vector3i, ByVal psTarget As Long, ByVal bVisible As Boolean, ByVal weapon_slot As Long) As Long
    proj_SendProjectileInternal = proj_SendProjectileAngledInternal(psWeap, psAttacker, player, target, psTarget, bVisible, weapon_slot, 0, FromU32(U32(gameTime) - 1))
End Function

' Send a single projectile against the given target.
' (VB) target must not be an aProj element (the pool may grow)
Public Function proj_SendProjectile(psWeap As WEAPON, ByVal psAttacker As Long, ByVal player As Long, target As Vector3i, ByVal psTarget As Long, ByVal bVisible As Boolean, ByVal weapon_slot As Long) As Boolean
    Dim p As Long
    p = proj_SendProjectileInternal(psWeap, psAttacker, player, target, psTarget, bVisible, weapon_slot)
    If p = 0 Then
        proj_SendProjectile = False
        Exit Function
    End If
    proj_AddActiveProjectile p
    proj_SendProjectile = True
End Function

' Send a single projectile against the given target with a minimum shot angle.
Public Function proj_SendProjectileAngled(psWeap As WEAPON, ByVal psAttacker As Long, ByVal player As Long, target As Vector3i, ByVal psTarget As Long, ByVal bVisible As Boolean, ByVal weapon_slot As Long, ByVal min_angle As Long, ByVal fireTime As Long) As Boolean
    Dim p As Long
    p = proj_SendProjectileAngledInternal(psWeap, psAttacker, player, target, psTarget, bVisible, weapon_slot, min_angle, fireTime)
    If p = 0 Then
        proj_SendProjectileAngled = False
        Exit Function
    End If
    proj_AddActiveProjectile p
    proj_SendProjectileAngled = True
End Function

Private Function INTERVAL_(ByVal b As Long, ByVal e As Long) As INTERVAL
    INTERVAL_.begin_ = b
    INTERVAL_.end_ = e
End Function

Private Function intervalIntersection(i1 As INTERVAL, i2 As INTERVAL) As INTERVAL
    intervalIntersection = INTERVAL_(MAX_(i1.begin_, i2.begin_), MIN_(i1.end_, i2.end_))
End Function

Private Function intervalEmpty(i As INTERVAL) As Boolean
    intervalEmpty = i.begin_ >= i.end_
End Function

Private Function collisionZ(ByVal z1 As Long, ByVal z2 As Long, ByVal height As Long) As INTERVAL
    Dim ret As INTERVAL
    ret = INTERVAL_(-1, -1)
    If z1 > z2 Then
        z1 = FromU32(-CDbl(z1))
        z2 = FromU32(-CDbl(z2))
    End If

    If z1 > height Or z2 < -height Then
        collisionZ = ret    ' No collision between time 1 and time 2.
        Exit Function
    End If

    If z1 = z2 Then
        If z1 >= -height And z1 <= height Then
            ret.begin_ = 0
            ret.end_ = 1024
        End If
        collisionZ = ret
        Exit Function
    End If

    ret.begin_ = FromU32(1024# * FromU32(-CDbl(height) - z1)) \ FromU32(CDbl(z2) - z1)
    ret.end_ = FromU32(1024# * FromU32(CDbl(height) - z1)) \ FromU32(CDbl(z2) - z1)
    collisionZ = ret
End Function

Private Function collisionXY(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal radius As Long) As INTERVAL
    Dim dx As Long, dy As Long, a As Variant, b As Variant, c As Variant, d As Variant, sd As Long
    Dim ret As INTERVAL

    Dim small As Boolean

    ' Solve (1 - t)v1 + t v2 = r.
    dx = FromU32(CDbl(x2) - x1)
    dy = FromU32(CDbl(y2) - y1)
    ' (VB) int64_t: Doubles when every operand is below 2^12 (a, b and c stay below 2^25, b^2 and a c below 2^50, all
    ' exact), else Decimals
    small = Abs(CDbl(x1)) < 4096 And Abs(CDbl(y1)) < 4096 And Abs(CDbl(dx)) < 4096 And Abs(CDbl(dy)) < 4096 And Abs(CDbl(radius)) < 4096
    If small Then
        a = CDbl(dx) * dx + CDbl(dy) * dy                   ' a = (v2 - v1)^2
        b = CDbl(x1) * dx + CDbl(y1) * dy                   ' b = v1(v2 - v1)
        c = CDbl(x1) * x1 + CDbl(y1) * y1 - CDbl(radius) * radius  ' c = v1^2 - r^2
    Else
        a = I64(CDec(dx) * dx + CDec(dy) * dy)
        b = I64(CDec(x1) * dx + CDec(y1) * dy)
        c = I64(CDec(x1) * x1 + CDec(y1) * y1 - CDec(radius) * radius)
    End If
    ' Equation to solve is now a t^2 + 2 b t + c = 0.
    If small Then d = b * b - a * c Else d = I64(b * b - a * c)  ' d = b^2 - a c
    ' Solution is (-b +- sqrt(d))/a.
    If d < 0 Then
        collisionXY = INTERVAL_(-1, -1)  ' Missed.
        Exit Function
    End If
    If a = 0 Then
        If c < 0 Then collisionXY = INTERVAL_(0, 1024) Else collisionXY = INTERVAL_(-1, -1)  ' Not moving. See if inside the target.
        Exit Function
    End If

    sd = i64Sqrt(d)
    If small Then
        ' (VB) below 2^38 over a: the Double quotient truncates right
        ret.begin_ = FromU32(MAX_d(0, Fix(1024 * (-b - sd) / a)))
        ret.end_ = FromU32(MIN_d(1024, Fix(1024 * (-b + sd) / a)))
    Else
        ret.begin_ = FromU32Dec(MAX_d(0, DivTrunc(I64(1024 * (-b - sd)), a)))
        ret.end_ = FromU32Dec(MIN_d(1024, DivTrunc(I64(1024 * (-b + sd)), a)))
    End If
    collisionXY = ret
End Function

' (VB) MAX and MIN of int64_t values (Doubles or Decimals)
Private Function MAX_d(ByVal a As Variant, ByVal b As Variant) As Variant
    If a > b Then MAX_d = a Else MAX_d = b
End Function

Private Function MIN_d(ByVal a As Variant, ByVal b As Variant) As Variant
    If a < b Then MIN_d = a Else MIN_d = b
End Function

Private Function collisionXYZ(v1 As Vector3i, v2 As Vector3i, shape As ObjectShape, ByVal height As Long) As Long
    Dim i As INTERVAL
    i = collisionZ(v1.z, v2.z, height)
    If Not intervalEmpty(i) Then  ' Don't bother checking x and y unless z passes.
        If shape.isRectangular Then
            i = intervalIntersection(i, collisionZ(v1.x, v2.x, shape.size.x))
            If Not intervalEmpty(i) Then  ' Don't bother checking y unless x and z pass.
                i = intervalIntersection(i, collisionZ(v1.y, v2.y, shape.size.y))
            End If
        Else  ' Else is circular.
            i = intervalIntersection(i, collisionXY(v1.x, v1.y, v2.x, v2.y, ObjectShape_getRadius(shape)))
        End If

        If Not intervalEmpty(i) Then
            collisionXYZ = MAX_(0, i.begin_)
            Exit Function
        End If
    End If
    collisionXYZ = -1
End Function

Private Function proj_InFlightFunc(ByVal psProj As Long) As Long
    ' we want a delay between Las-Sats firing and actually hitting in multiPlayer
    ' magic number but that's how long the audio countdown message lasts!
    Const LAS_SAT_DELAY As Long = 4
    Dim closestCollisionObject As Long, closestCollisionSpacetime As Spacetime, timeSoFar As Long, deltaProjectileTime As Long
    Dim psStats As Long, currentDistance As Long, delta As Vector3i, targetDistance As Long, targetDroid As Long
    Dim flightTime As Long, r As Long, ang As Long, destPos As Vector3i, horizontalTargetDistance As Long
    Dim terrainHeight As Long, desiredMinHeight As Long, desiredMaxHeight As Long, heightError As Long
    Dim step_ As Vector3i, numer As Vector3i, tries As Long, n As Long, gi As Long, psTempObj As Long
    Dim psTempObjPrevPos As Vector3i, diff As Vector3i, prevDiff As Vector3i, targetHeight As Long, targetShape As ObjectShape
    Dim collision As Long, collisionTime As Long, terrainIntersectTime As Long, spawnedProjectile As Long
    Dim asWeap As WEAPON, dstCopy As Vector3i, effectTime As Long, st As Spacetime, posFlip As Vector3i
    Dim c As Boolean, objPos As Vector3i, p As Vector3i, longRange As Long

    closestCollisionObject = 0

    checkProjectile psProj, "src/projectile.cpp:751", "proj_InFlightFunc", max_check_object_recursion

    timeSoFar = FromU32(U32(gameTime) - U32(aProj(psProj).born))

    aProj(psProj).time_ = gameTime
    deltaProjectileTime = FromU32(U32(aProj(psProj).time_) - U32(aProj(psProj).prevSpacetime.time_))

    psStats = aProj(psProj).psWStats
    If Not (psStats <> -1) Then ASSERT_FAILURE "psStats != nullptr", "src/projectile.cpp:759", "proj_InFlightFunc", "Invalid weapon stats pointer": proj_InFlightFunc = 0: Exit Function

    ' we want a delay between Las-Sats firing and actually hitting in multiPlayer
    ' magic number but that's how long the audio countdown message lasts!
    If bMultiPlayer And asWeaponStats(psStats).weaponSubClass = WSC_LAS_SAT And _
       ULT(timeSoFar, LAS_SAT_DELAY * GAME_TICKS_PER_SEC) Then
        proj_InFlightFunc = 0
        Exit Function
    End If

    ' Calculate movement vector:
    currentDistance = 0
    Select Case asWeaponStats(psStats).movementModel
    Case MM_DIRECT           ' Go in a straight line.
        delta.x = aProj(psProj).dst.x - aProj(psProj).src.x
        delta.y = aProj(psProj).dst.y - aProj(psProj).src.y
        delta.z = aProj(psProj).dst.z - aProj(psProj).src.z
        If asWeaponStats(psStats).weaponSubClass = WSC_LAS_SAT Then
            ' LASSAT doesn't have a z
            delta.z = 0
        End If
        targetDistance = MAX_(iHypot(delta.x, delta.y), 1)
        currentDistance = FromU32(Int(U32(MUL32(timeSoFar, asWeaponStats(psStats).flightSpeed)) / GAME_TICKS_PER_SEC))   ' (unsigned)
        aProj(psProj).pos.x = aProj(psProj).src.x + MUL32(delta.x, currentDistance) \ targetDistance
        aProj(psProj).pos.y = aProj(psProj).src.y + MUL32(delta.y, currentDistance) \ targetDistance
        aProj(psProj).pos.z = aProj(psProj).src.z + MUL32(delta.z, currentDistance) \ targetDistance
    Case MM_INDIRECT         ' Ballistic trajectory.
        delta.x = aProj(psProj).dst.x - aProj(psProj).src.x
        delta.y = aProj(psProj).dst.y - aProj(psProj).src.y
        delta.z = MUL32(aProj(psProj).vZ - MUL32(timeSoFar, ACC_GRAVITY) \ (GAME_TICKS_PER_SEC * 2), timeSoFar) \ GAME_TICKS_PER_SEC ' '2' because we reach our highest point in the mid of flight, when "vZ is 0".
        targetDistance = MAX_(iHypot(delta.x, delta.y), 1)
        currentDistance = MUL32(timeSoFar, aProj(psProj).vXY) \ GAME_TICKS_PER_SEC
        aProj(psProj).pos.x = aProj(psProj).src.x + MUL32(delta.x, currentDistance) \ targetDistance
        aProj(psProj).pos.y = aProj(psProj).src.y + MUL32(delta.y, currentDistance) \ targetDistance
        aProj(psProj).pos.z = aProj(psProj).src.z + MUL32(delta.z, currentDistance) \ targetDistance
        aProj(psProj).pos.z = aProj(psProj).src.z + delta.z  ' Use raw z value.
        aProj(psProj).rot.pitch = iAtan2(aProj(psProj).vZ - MUL32(timeSoFar, ACC_GRAVITY) \ GAME_TICKS_PER_SEC, aProj(psProj).vXY)
    Case MM_HOMINGDIRECT, MM_HOMINGINDIRECT   ' Fly towards target, even if target moves. (MM_HOMINGINDIRECT: Avoid terrain.)
        If aProj(psProj).psDest <> 0 Then
            destPos = SIMPLE_OBJECT_pos(aProj(psProj).psDest)
            If asWeaponStats(psStats).movementModel = MM_HOMINGDIRECT Then
                ' If it's homing and has a target (not a miss)...
                ' Home at the centre of the part that was visible when firing.
                aProj(psProj).dst = Vector3i_(destPos.x, destPos.y, destPos.z + (establishTargetHeight(aProj(psProj).psDest) - aProj(psProj).partVisible \ 2))
            Else
                aProj(psProj).dst = Vector3i_(destPos.x, destPos.y, destPos.z + establishTargetHeight(aProj(psProj).psDest) \ 2)
            End If
            targetDroid = castDroid(aProj(psProj).psDest)
            If targetDroid <> 0 Then
                ' Do target prediction.
                delta.x = aProj(psProj).dst.x - aProj(psProj).pos.x
                delta.y = aProj(psProj).dst.y - aProj(psProj).pos.y
                delta.z = aProj(psProj).dst.z - aProj(psProj).pos.z
                flightTime = FromU32(Fix(U32(MUL32(iHypot(delta.x, delta.y), GAME_TICKS_PER_SEC)) / U32(asWeaponStats(psStats).flightSpeed)))
                r = MUL32(MIN_(aDroid(targetDroid).sMove.speed, FromU32(Int(U32(MUL32(asWeaponStats(psStats).flightSpeed, 3)) / 4))), flightTime) \ GAME_TICKS_PER_SEC
                aProj(psProj).dst.x = aProj(psProj).dst.x + iSinR(aDroid(targetDroid).sMove.moveDir, r)
                aProj(psProj).dst.y = aProj(psProj).dst.y + iCosR(aDroid(targetDroid).sMove.moveDir, r)
            End If
            aProj(psProj).dst.x = clip_(aProj(psProj).dst.x, 0, world_coord(mapWidth) - 1)
            aProj(psProj).dst.y = clip_(aProj(psProj).dst.y, 0, world_coord(mapHeight) - 1)
        End If
        If asWeaponStats(psStats).movementModel = MM_HOMINGINDIRECT Then
            If aProj(psProj).psDest = 0 Then
                aProj(psProj).dst.z = map_Height(aProj(psProj).pos.x, aProj(psProj).pos.y) - 1  ' Target missing, so just home in on the ground under where the target was.
            End If
            horizontalTargetDistance = iHypot(aProj(psProj).dst.x - aProj(psProj).pos.x, aProj(psProj).dst.y - aProj(psProj).pos.y)
            ang = iAtan2(aProj(psProj).dst.x - aProj(psProj).pos.x, aProj(psProj).dst.y - aProj(psProj).pos.y)
            r = FromU32(Fix(U32(MUL32(MUL32(asWeaponStats(psStats).flightSpeed, 2), deltaProjectileTime)) / GAME_TICKS_PER_SEC))
            terrainHeight = MAX_(map_Height(aProj(psProj).pos.x, aProj(psProj).pos.y), _
                                 map_Height(aProj(psProj).pos.x + iSinR(ang, r), aProj(psProj).pos.y + iCosR(ang, r)))
            desiredMinHeight = terrainHeight + MIN_(horizontalTargetDistance \ 4, HOMINGINDIRECT_HEIGHT_MIN)
            desiredMaxHeight = MAX_(aProj(psProj).dst.z, terrainHeight + HOMINGINDIRECT_HEIGHT_MAX)
            heightError = aProj(psProj).pos.z - clip_(aProj(psProj).pos.z, desiredMinHeight, desiredMaxHeight)
            aProj(psProj).dst.z = FromU32(CDbl(aProj(psProj).dst.z) - (MUL32(MUL32(horizontalTargetDistance, heightError), 2) \ HOMINGINDIRECT_HEIGHT_MIN))
        End If
        delta.x = aProj(psProj).dst.x - aProj(psProj).pos.x
        delta.y = aProj(psProj).dst.y - aProj(psProj).pos.y
        delta.z = aProj(psProj).dst.z - aProj(psProj).pos.z
        targetDistance = MAX_(iHypot3(delta.x, delta.y, delta.z), 1)
        If aProj(psProj).psDest = 0 And targetDistance < 10000 And asWeaponStats(psStats).movementModel = MM_HOMINGDIRECT Then
            aProj(psProj).dst = Vector3i_(aProj(psProj).pos.x + MUL32(delta.x, 10), aProj(psProj).pos.y + MUL32(delta.y, 10), aProj(psProj).pos.z + MUL32(delta.z, 10)) ' Target missing, so just keep going in a straight line.
        End If
        currentDistance = FromU32(Int(U32(MUL32(timeSoFar, asWeaponStats(psStats).flightSpeed)) / GAME_TICKS_PER_SEC))   ' (unsigned)
        numer = Vector3i_(MUL32(delta.x, asWeaponStats(psStats).flightSpeed), MUL32(delta.y, asWeaponStats(psStats).flightSpeed), MUL32(delta.z, asWeaponStats(psStats).flightSpeed))
        step_ = quantiseFraction_vec(numer, MUL32(GAME_TICKS_PER_SEC, targetDistance), aProj(psProj).time_, aProj(psProj).prevSpacetime.time_)
        If asWeaponStats(psStats).movementModel = MM_HOMINGINDIRECT And aProj(psProj).psDest <> 0 Then
            tries = 0
            Do While tries < 10
                p = Vector3i_(aProj(psProj).pos.x + step_.x, aProj(psProj).pos.y + step_.y, aProj(psProj).pos.z + step_.z)
                If Not ULT(map_LineIntersect(aProj(psProj).prevSpacetime.pos.x, aProj(psProj).prevSpacetime.pos.y, aProj(psProj).prevSpacetime.pos.z, _
                                             p.x, p.y, p.z, iHypot3(step_.x, step_.y, step_.z)), FromU32(U32(targetDistance) - 1)) Then Exit Do
                aProj(psProj).dst.z = aProj(psProj).dst.z + iHypot(aProj(psProj).dst.x - aProj(psProj).pos.x, aProj(psProj).dst.y - aProj(psProj).pos.y)  ' Would collide with terrain this tick, change trajectory.
                ' Recalculate delta, targetDistance and step.
                delta.x = aProj(psProj).dst.x - aProj(psProj).pos.x
                delta.y = aProj(psProj).dst.y - aProj(psProj).pos.y
                delta.z = aProj(psProj).dst.z - aProj(psProj).pos.z
                targetDistance = MAX_(iHypot3(delta.x, delta.y, delta.z), 1)
                numer = Vector3i_(MUL32(delta.x, asWeaponStats(psStats).flightSpeed), MUL32(delta.y, asWeaponStats(psStats).flightSpeed), MUL32(delta.z, asWeaponStats(psStats).flightSpeed))
                step_ = quantiseFraction_vec(numer, MUL32(GAME_TICKS_PER_SEC, targetDistance), aProj(psProj).time_, aProj(psProj).prevSpacetime.time_)
                tries = tries + 1
            Loop
        End If
        aProj(psProj).pos.x = aProj(psProj).pos.x + step_.x
        aProj(psProj).pos.y = aProj(psProj).pos.y + step_.y
        aProj(psProj).pos.z = aProj(psProj).pos.z + step_.z
        aProj(psProj).rot.direction = iAtan2(delta.x, delta.y)
        aProj(psProj).rot.pitch = iAtan2(delta.z, targetDistance)
    End Select

    closestCollisionSpacetime.time_ = &HFFFFFFFF

    ' Check nearby objects for possible collisions
    n = gridStartIterate(aProj(psProj).pos.x, aProj(psProj).pos.y, PROJ_NEIGHBOUR_RANGE, gridListInFlight)
    For gi = 0 To n - 1
        psTempObj = gridListInFlight(gi)
        checkObject psTempObj, "src/projectile.cpp:870", "proj_InFlightFunc", max_check_object_recursion

        ' (VB) each branch of the else-if chain continues: one If each
        If psDamaged_contains(psProj, psTempObj) Then
            ' Dont damage one target twice
            GoTo continue_
        End If
        If SIMPLE_OBJECT_died(psTempObj) <> 0 Then
            ' Do not damage dead objects further
            If Not (SIMPLE_OBJECT_type(psTempObj) < OBJ_NUM_TYPES) Then ASSERT_FAILURE "psTempObj->type < OBJ_NUM_TYPES", "src/projectile.cpp:880", "proj_InFlightFunc", "Bad pointer! type=%u", SIMPLE_OBJECT_type(psTempObj)
            GoTo continue_
        End If
        If SIMPLE_OBJECT_type(psTempObj) = OBJ_FEATURE Then
            If Not asFeatureStats(aFeature(psTempObj).psStats).damageable Then
                ' Ignore oil resources, artifacts and other pickups
                GoTo continue_
            End If
        End If
        If aiCheckAlliances(SIMPLE_OBJECT_player(psTempObj), aProj(psProj).player) And psTempObj <> aProj(psProj).psDest Then
            ' No friendly fire unless intentional
            GoTo continue_
        End If
        If (asWeaponStats(psStats).surfaceToAir And SHOOT_ON_GROUND) = 0 Then
            c = SIMPLE_OBJECT_type(psTempObj) = OBJ_STRUCTURE Or SIMPLE_OBJECT_type(psTempObj) = OBJ_FEATURE
            If Not c And SIMPLE_OBJECT_type(psTempObj) = OBJ_DROID Then
                If Not DROID_isFlightBasedTransporter(psTempObj) Then c = Not DROID_isFlying(psTempObj)
            End If
            If c Then
                ' AA weapons should not hit buildings and non-vtol droids
                GoTo continue_
            End If
        End If

        objPos = SIMPLE_OBJECT_pos(psTempObj)
        If isDroid(psTempObj) Then psTempObjPrevPos = aDroid(psTempObj).prevSpacetime.pos Else psTempObjPrevPos = objPos

        diff.x = aProj(psProj).pos.x - objPos.x
        diff.y = aProj(psProj).pos.y - objPos.y
        diff.z = aProj(psProj).pos.z - objPos.z
        prevDiff.x = aProj(psProj).prevSpacetime.pos.x - psTempObjPrevPos.x
        prevDiff.y = aProj(psProj).prevSpacetime.pos.y - psTempObjPrevPos.y
        prevDiff.z = aProj(psProj).prevSpacetime.pos.z - psTempObjPrevPos.z
        targetHeight = establishTargetHeight(psTempObj)
        targetShape = establishTargetShape(psTempObj)
        collision = collisionXYZ(prevDiff, diff, targetShape, targetHeight)
        collisionTime = FromU32(U32(aProj(psProj).prevSpacetime.time_) + _
                                Int(U32(MUL32(FromU32(U32(aProj(psProj).time_) - U32(aProj(psProj).prevSpacetime.time_)), collision)) / 1024))

        If collision >= 0 And ULT(collisionTime, closestCollisionSpacetime.time_) Then
            ' We hit!
            closestCollisionSpacetime = interpolateObjectSpacetime(psProj, collisionTime)
            closestCollisionObject = psTempObj

            ' Keep testing for more collisions, in case there was a closer target.
        End If
continue_:
    Next

    terrainIntersectTime = map_LineIntersect(aProj(psProj).prevSpacetime.pos.x, aProj(psProj).prevSpacetime.pos.y, aProj(psProj).prevSpacetime.pos.z, _
                                             aProj(psProj).pos.x, aProj(psProj).pos.y, aProj(psProj).pos.z, _
                                             FromU32(U32(aProj(psProj).time_) - U32(aProj(psProj).prevSpacetime.time_)))
    If terrainIntersectTime <> &HFFFFFFFF Then
        collisionTime = FromU32(U32(aProj(psProj).prevSpacetime.time_) + U32(terrainIntersectTime))
        If ULT(collisionTime, closestCollisionSpacetime.time_) Then
            ' We hit the terrain!
            closestCollisionSpacetime = interpolateObjectSpacetime(psProj, collisionTime)
            closestCollisionObject = 0
        End If
    End If

    If closestCollisionSpacetime.time_ <> &HFFFFFFFF Then
        ' We hit!
        setSpacetime psProj, closestCollisionSpacetime
        If ULT(aProj(psProj).time_, FromU32(U32(gameTime) - U32(deltaGameTime) + 1)) Then aProj(psProj).time_ = FromU32(U32(gameTime) - U32(deltaGameTime) + 1)  ' Make sure .died gets set in the interval [gameTime - deltaGameTime + 1; gameTime].
        If aProj(psProj).time_ = aProj(psProj).prevSpacetime.time_ Then
            aProj(psProj).prevSpacetime.time_ = FromU32(U32(aProj(psProj).prevSpacetime.time_) - 1)
        End If
        setProjectileDestination psProj, closestCollisionObject  ' We hit something.

        spawnedProjectile = 0
        ' Buildings and terrain cannot be penetrated and we need a penetrating weapon, and projectile should not have already travelled further than 1.25 * maximum range.
        c = False
        If closestCollisionObject <> 0 Then
            If SIMPLE_OBJECT_type(closestCollisionObject) = OBJ_DROID And asWeaponStats(psStats).penetrate Then
                c = currentDistance < Fix(1.25 * proj_GetLongRange(asWeaponStats(psStats), aProj(psProj).player))
            End If
        End If
        If c Then
            asWeap.nStat = psStats

            ' Assume we damaged the chosen target
            psDamaged_push psProj, closestCollisionObject

            dstCopy = aProj(psProj).dst
            spawnedProjectile = proj_SendProjectileInternal(asWeap, psProj, aProj(psProj).player, dstCopy, 0, True, -1)
        End If

        aProj(psProj).state = PROJ_IMPACT

        proj_InFlightFunc = spawnedProjectile
        Exit Function
    End If

    longRange = proj_GetLongRange(asWeaponStats(psStats), aProj(psProj).player)
    If MUL32(currentDistance, 100) >= MUL32(longRange, asWeaponStats(psStats).distanceExtensionFactor) Then
        ' We've travelled our maximum range.
        aProj(psProj).state = PROJ_IMPACT
        setProjectileDestination psProj, 0 ' miss registered if NULL target
        proj_InFlightFunc = 0
        Exit Function
    End If

    ' Paint effects if visible
    If gfxVisible(psProj) Then
        effectTime = FromU32(U32(aProj(psProj).prevSpacetime.time_) + 31) And Not 31&
        Do While ULT(effectTime, aProj(psProj).time_)
            st = interpolateObjectSpacetime(psProj, effectTime)
            posFlip = Vector3i_(st.pos.x, st.pos.z, st.pos.y)
            Select Case asWeaponStats(psStats).weaponSubClass
            Case WSC_FLAME
                posFlip.z = posFlip.z - 8  ' Why?
                effectGiveAuxVar MUL32(currentDistance, 100) \ proj_GetLongRange(asWeaponStats(psStats), aProj(psProj).player)
                addEffect posFlip, EFFECT_EXPLOSION, EXPLOSION_TYPE_FLAMETHROWER, False, 0, 0, effectTime
            Case WSC_COMMAND, WSC_ELECTRONIC, WSC_EMP
                posFlip.z = posFlip.z - 8  ' Why?
                effectGiveAuxVar (MUL32(currentDistance, 100) \ proj_GetLongRange(asWeaponStats(psStats), aProj(psProj).player)) \ 2
                addEffect posFlip, EFFECT_EXPLOSION, EXPLOSION_TYPE_LASER, False, 0, 0, effectTime
            Case WSC_ROCKET, WSC_MISSILE, WSC_SLOWROCKET, WSC_SLOWMISSILE
                posFlip.z = posFlip.z + 8  ' Why?
                addEffect posFlip, EFFECT_SMOKE, SMOKE_TYPE_TRAIL, False, 0, 0, effectTime
            Case Else
                ' Add smoke trail to indirect weapons, even if firing directly.
                If Not proj_Direct(psStats) Then
                    posFlip.z = posFlip.z + 4  ' Why?
                    addEffect posFlip, EFFECT_SMOKE, SMOKE_TYPE_TRAIL, False, 0, 0, effectTime
                End If
                ' Otherwise no effect.
            End Select
            effectTime = FromU32(U32(effectTime) + 32)
        Loop
    End If
    proj_InFlightFunc = 0
End Function

' targetPos: not an aProj element
Private Sub proj_radiusSweep(ByVal psObj As Long, ByVal psStats As Long, targetPos As Vector3i, ByVal empRadius As Boolean)
    Dim n As Long, gi As Long, psCurr As Long, bTargetInAir As Boolean, useSphere As Boolean, damageable As Boolean
    Dim targetInFlag As Long, damage As Long, radius As Long, sDamage As DAMAGE_, objPos As Vector3i, player As Long

    player = aProj(psObj).player
    If empRadius Then radius = asWeaponStats(psStats).upgrade(player).empRadius Else radius = asWeaponStats(psStats).upgrade(player).radius
    n = gridStartIterate(targetPos.x, targetPos.y, radius, gridListSweep)

    For gi = 0 To n - 1
        psCurr = gridListSweep(gi)
        If SIMPLE_OBJECT_died(psCurr) <> 0 Then
            If Not (SIMPLE_OBJECT_type(psCurr) < OBJ_NUM_TYPES) Then ASSERT_FAILURE "psCurr->type < OBJ_NUM_TYPES", "src/projectile.cpp:1027", "proj_radiusSweep", "Bad pointer! type=%u", SIMPLE_OBJECT_type(psCurr)
            GoTo continue_  ' Do not damage dead objects further.
        End If

        If psCurr = aProj(psObj).psDest Then
            GoTo continue_  ' Don't hit main target twice.
        End If

        If aProj(psObj).psSource <> 0 Then
            If SIMPLE_OBJECT_player(aProj(psObj).psSource) = SIMPLE_OBJECT_player(psCurr) And (asWeaponStats(psStats).flags And BIT_(WEAPON_FLAG_NO_FRIENDLY_FIRE)) <> 0 Then
                GoTo continue_ ' this weapon does not do friendly damage
            End If
        End If

        bTargetInAir = False
        useSphere = False
        damageable = True
        Select Case SIMPLE_OBJECT_type(psCurr)
        Case OBJ_DROID
            bTargetInAir = asPropulsionTypes(asPropulsionStats(DROID_getPropulsionStats(psCurr)).propulsionType).travel = AIR And aDroid(psCurr).sMove.Status <> MOVEINACTIVE
            useSphere = True
        Case OBJ_STRUCTURE
        Case OBJ_FEATURE
            damageable = asFeatureStats(aFeature(psCurr).psStats).damageable
        Case Else
            ASSERT_FAILURE "false", "src/projectile.cpp:1055", "proj_radiusSweep", "Bad type."
            GoTo continue_
        End Select

        If Not damageable Then
            GoTo continue_  ' Ignore features that are not damageable.
        End If
        If bTargetInAir Then targetInFlag = SHOOT_IN_AIR Else targetInFlag = SHOOT_ON_GROUND
        If (asWeaponStats(psStats).surfaceToAir And targetInFlag) = 0 Then
            GoTo continue_  ' Target in air, and can't shoot at air, or target on ground, and can't shoot at ground.
        End If
        If useSphere Then
            objPos = SIMPLE_OBJECT_pos(psCurr)
            If Not Vector3i_InSphere(objPos, targetPos, radius) Then
                GoTo continue_  ' Target out of range.
            End If
        End If
        ' The psCurr will get damaged, at this point.
        damage = calcDamage(weaponRadDamage(asWeaponStats(psStats), player), asWeaponStats(psStats).weaponEffect, psCurr)
        debug_ 1073, LOG_ATTACK, "proj_radiusSweep", "Damage to object %d, player %d : %u", SIMPLE_OBJECT_id(psCurr), SIMPLE_OBJECT_player(psCurr), damage
        If bMultiPlayer And aProj(psObj).psSource <> 0 And SIMPLE_OBJECT_type(psCurr) <> OBJ_FEATURE Then
            updateMultiStatsDamage SIMPLE_OBJECT_player(aProj(psObj).psSource), SIMPLE_OBJECT_player(psCurr), damage
        End If

        sDamage.psProjectile = psObj
        sDamage.psDest = psCurr
        sDamage.damage = damage
        sDamage.weaponClass = asWeaponStats(psStats).weaponClass
        sDamage.weaponSubClass = asWeaponStats(psStats).weaponSubClass
        sDamage.impactTime = aProj(psObj).time_
        sDamage.isDamagePerSecond = False
        sDamage.minDamage = asWeaponStats(psStats).upgrade(player).minimumDamage
        sDamage.empRadiusHit = empRadius

        objectDamage sDamage
continue_:
    Next
End Sub

Private Sub proj_ImpactFunc(ByVal psObj As Long)
    Dim psStats As Long, iAudioImpactID As Long, relativeDamage As Long, position As Vector3i, scatter As Vector3i
    Dim imd As Long, temp As Long, hasRadius As Boolean, hasEMPRadius As Boolean, facing As EFFECT_TYPE
    Dim damage As Long, sDamage As DAMAGE_, destDroid As Long, targetPos As Vector3i, destPos As Vector3i
    Dim player As Long, displayImd As Long

    If Not (psObj <> 0) Then ASSERT_FAILURE "psObj != nullptr", "src/projectile.cpp:1107", "proj_ImpactFunc", "Invalid pointer": Exit Sub
    checkProjectile psObj, "src/projectile.cpp:1108", "proj_ImpactFunc", max_check_object_recursion

    psStats = aProj(psObj).psWStats
    If Not (psStats <> -1) Then ASSERT_FAILURE "psStats != nullptr", "src/projectile.cpp:1111", "proj_ImpactFunc", "Invalid weapon stats pointer": Exit Sub
    player = aProj(psObj).player

    ' play impact audio
    If gfxVisible(psObj) Then
        If asWeaponStats(psStats).iAudioImpactID = NO_SOUND Then
            ' play richochet if MG
            If aProj(psObj).psDest <> 0 And asWeaponStats(psStats).weaponSubClass = WSC_MGUN Then
                If rand() Mod 3 = 0 Then    ' ONEINTHREE
                    iAudioImpactID = ID_SOUND_RICOCHET_1 + (rand() Mod 3)
                    destPos = SIMPLE_OBJECT_pos(aProj(psObj).psDest)
                    audio_PlayStaticTrack destPos.x, destPos.y, iAudioImpactID
                End If
            End If
        Else
            audio_PlayStaticTrack aProj(psObj).pos.x, aProj(psObj).pos.y, asWeaponStats(psStats).iAudioImpactID
        End If

        ' Shouldn't need to do this check but the stats aren't all at a value yet... FIXME
        If asWeaponStats(psStats).upgrade(player).periodicalDamageRadius <> 0 And asWeaponStats(psStats).upgrade(player).periodicalDamageTime <> 0 Then
            position.x = aProj(psObj).pos.x
            position.z = aProj(psObj).pos.y ' z = y [sic] intentional
            position.y = map_Height(position.x, position.z)
            effectGiveAuxVar asWeaponStats(psStats).upgrade(player).periodicalDamageRadius
            effectGiveAuxVarSec asWeaponStats(psStats).upgrade(player).periodicalDamageTime
            addEffect position, EFFECT_FIRE, FIRE_TYPE_LOCALISED, False, 0, 0, aProj(psObj).time_
        End If

        ' may want to add both a fire effect and the las sat effect
        If asWeaponStats(psStats).weaponSubClass = WSC_LAS_SAT Then
            position.x = aProj(psObj).pos.x
            position.z = aProj(psObj).pos.y  ' z = y [sic] intentional
            position.y = map_Height(position.x, position.z)
            addEffect position, EFFECT_SAT_LASER, SAT_LASER_STANDARD, False, 0, 0, aProj(psObj).time_

            If clipXY(aProj(psObj).pos.x, aProj(psObj).pos.y) Then
                shakeStart 1800 ' takes out lots of stuff so shake length is greater
            End If
        End If
    End If

    If asWeaponStats(psStats).upgrade(player).periodicalDamageRadius <> 0 And asWeaponStats(psStats).upgrade(player).periodicalDamageTime <> 0 Then
        tileSetFire aProj(psObj).pos.x, aProj(psObj).pos.y, asWeaponStats(psStats).upgrade(player).periodicalDamageTime
    End If

    ' Set the effects position and radius
    position.x = aProj(psObj).pos.x
    position.z = aProj(psObj).pos.y ' z = y [sic] intentional
    position.y = aProj(psObj).pos.z ' y = z [sic] intentional
    scatter.x = asWeaponStats(psStats).upgrade(player).radius
    scatter.y = 0
    scatter.z = asWeaponStats(psStats).upgrade(player).radius

    ' If the projectile missed its target (or the target died)
    If aProj(psObj).psDest = 0 Then
        If gfxVisible(psObj) Then
            ' Get if we are facing or not
            If asWeaponStats(psStats).facePlayer <> 0 Then facing = EXPLOSION_TYPE_SPECIFIED Else facing = EXPLOSION_TYPE_NOT_FACING

            ' The graphic to show depends on if we hit water or not
            If terrainType(mapTile(map_coord(aProj(psObj).pos.x), map_coord(aProj(psObj).pos.y))) = TER_WATER Then
                imd = asWeaponStats(psStats).pWaterHitGraphic
            ' We did not hit water, the regular miss graphic will do the trick
            Else
                imd = asWeaponStats(psStats).pTargetMissGraphic
            End If

            If imd <> 0 Then displayImd = iIMDBaseShape_displayModel(imd) Else displayImd = 0
            addMultiEffect position, scatter, EFFECT_EXPLOSION, facing, True, displayImd, asWeaponStats(psStats).numExplosions, asWeaponStats(psStats).lightWorld, asWeaponStats(psStats).effectSize, aProj(psObj).time_

            ' If the target was a VTOL hit in the air add smoke
            If (asWeaponStats(psStats).surfaceToAir And SHOOT_IN_AIR) <> 0 _
               And (asWeaponStats(psStats).surfaceToAir And SHOOT_ON_GROUND) = 0 Then
                addMultiEffect position, scatter, EFFECT_SMOKE, SMOKE_TYPE_DRIFTING, False, 0, 3, False, 0, aProj(psObj).time_
            End If
        End If
    ' The projectile hit its intended target
    Else
        checkObject aProj(psObj).psDest, "src/projectile.cpp:1202", "proj_ImpactFunc", max_check_object_recursion

        If SIMPLE_OBJECT_type(aProj(psObj).psDest) = OBJ_FEATURE Then
            If Not asFeatureStats(aFeature(aProj(psObj).psDest).psStats).damageable Then
                debug_ 1207, LOG_NEVER, "proj_ImpactFunc", "proj_ImpactFunc: trying to damage non-damageable target,projectile removed"
                aProj(psObj).state = PROJ_INACTIVE
                Exit Sub
            End If
        End If

        If gfxVisible(psObj) Then
            ' Get if we are facing or not
            If asWeaponStats(psStats).facePlayer <> 0 Then facing = EXPLOSION_TYPE_SPECIFIED Else facing = EXPLOSION_TYPE_NOT_FACING

            ' If we hit a VTOL with an AA gun use the miss graphic and add some smoke
            If (asWeaponStats(psStats).surfaceToAir And SHOOT_IN_AIR) <> 0 _
               And (asWeaponStats(psStats).surfaceToAir And SHOOT_ON_GROUND) = 0 _
               And asWeaponStats(psStats).weaponSubClass = WSC_AAGUN Then
                imd = asWeaponStats(psStats).pTargetMissGraphic
                addMultiEffect position, scatter, EFFECT_SMOKE, SMOKE_TYPE_DRIFTING, False, 0, 3, False, 0, aProj(psObj).time_
            ' Otherwise we just hit it plain and simple
            Else
                imd = asWeaponStats(psStats).pTargetHitGraphic
            End If

            If imd <> 0 Then displayImd = iIMDBaseShape_displayModel(imd) Else displayImd = 0
            addMultiEffect position, scatter, EFFECT_EXPLOSION, facing, True, displayImd, asWeaponStats(psStats).numExplosions, asWeaponStats(psStats).lightWorld, asWeaponStats(psStats).effectSize, aProj(psObj).time_
        End If

        ' Check for electronic warfare damage where we know the subclass and source
        If asWeaponStats(psStats).weaponSubClass = WSC_ELECTRONIC _
           And aProj(psObj).psSource <> 0 Then
            ' If we did enough `damage' to capture the target
            If electronicDamage(aProj(psObj).psDest, aProj(psObj).psSource, _
                                calcDamage(weaponDamage(asWeaponStats(psStats), player), asWeaponStats(psStats).weaponEffect, aProj(psObj).psDest), _
                                player, (asWeaponStats(psStats).flags And BIT_(WEAPON_FLAG_TELEPORT_CAPTURE)) <> 0) Then
                Select Case SIMPLE_OBJECT_type(aProj(psObj).psSource)
                Case OBJ_DROID
                    aDroid(aProj(psObj).psSource).order.typ = DORDER_NONE
                    actionDroid aProj(psObj).psSource, DACTION_NONE

                Case OBJ_STRUCTURE
                    aStruct(aProj(psObj).psSource).psTarget(0) = 0

                ' This is only here to prevent the compiler from producing
                ' warnings for unhandled enumeration values
                Case Else
                End Select
            End If
        ' Else it is just a regular weapon (direct or indirect)
        Else
            ' Calculate the damage the weapon does to its target
            damage = calcDamage(weaponDamage(asWeaponStats(psStats), player), asWeaponStats(psStats).weaponEffect, aProj(psObj).psDest)

            ' If we are in a multi-player game and the attacker is our responsibility
            If bMultiPlayer And aProj(psObj).psSource <> 0 Then
                updateMultiStatsDamage SIMPLE_OBJECT_player(aProj(psObj).psSource), SIMPLE_OBJECT_player(aProj(psObj).psDest), damage
            End If

            debug_ 1273, LOG_NEVER, "proj_ImpactFunc", "Damage to object %d, player %d\n", _
                   SIMPLE_OBJECT_id(aProj(psObj).psDest), SIMPLE_OBJECT_player(aProj(psObj).psDest)

            sDamage.psProjectile = psObj
            sDamage.psDest = aProj(psObj).psDest
            sDamage.damage = damage
            sDamage.weaponClass = asWeaponStats(psStats).weaponClass
            sDamage.weaponSubClass = asWeaponStats(psStats).weaponSubClass
            sDamage.impactTime = aProj(psObj).time_
            sDamage.isDamagePerSecond = False
            sDamage.minDamage = asWeaponStats(psStats).upgrade(player).minimumDamage
            sDamage.empRadiusHit = False

            ' Damage the object
            relativeDamage = objectDamage(sDamage)

            If relativeDamage >= 0 Then ' So long as the target wasn't killed
                psDamaged_push psObj, aProj(psObj).psDest
            End If
        End If
    End If

    temp = aProj(psObj).psDest
    setProjectileDestination psObj, 0
    ' The damage has been done, no more damage expected from this projectile. (Ignore periodical damaging.)
    aProj(psObj).expectedDamageCaused = 0
    setProjectileDestination psObj, temp

    ' If the projectile does no splash damage and does not set fire to things
    If asWeaponStats(psStats).upgrade(player).radius = 0 And asWeaponStats(psStats).upgrade(player).periodicalDamageTime = 0 Then
        aProj(psObj).state = PROJ_INACTIVE
        Exit Sub
    End If

    hasRadius = asWeaponStats(psStats).upgrade(player).radius <> 0
    hasEMPRadius = asWeaponStats(psStats).upgrade(player).empRadius <> 0
    If hasRadius Or hasEMPRadius Then
        ' An area effect bullet
        aProj(psObj).state = PROJ_POSTIMPACT

        ' Note when it exploded for the explosion effect
        aProj(psObj).born = gameTime

        ' If projectile impacts a droid start the splash damage from the center of it, else use whatever location the projectile impacts at as the splash center.
        destDroid = castDroid(aProj(psObj).psDest)
        If destDroid <> 0 Then targetPos = aDroid(destDroid).pos Else targetPos = aProj(psObj).pos

        If hasEMPRadius And asWeaponStats(psStats).weaponSubClass = WSC_EMP Then
            proj_radiusSweep psObj, psStats, targetPos, True
        End If
        If hasRadius Then
            proj_radiusSweep psObj, psStats, targetPos, False
        End If
    End If

    If asWeaponStats(psStats).upgrade(player).periodicalDamageTime <> 0 Then
        ' Periodical damage round
        ' Periodical damage gets done in the bullet update routine
        ' Just note when started damaging
        aProj(psObj).state = PROJ_POSTIMPACT
        aProj(psObj).born = gameTime
    End If
    ' Something was blown up
End Sub

Private Sub proj_PostImpactFunc(ByVal psObj As Long)
    Dim psStats As Long, age As Long

    If Not (psObj <> 0) Then ASSERT_FAILURE "psObj != nullptr", "src/projectile.cpp:1350", "proj_PostImpactFunc", "Invalid pointer": Exit Sub
    checkProjectile psObj, "src/projectile.cpp:1351", "proj_PostImpactFunc", max_check_object_recursion

    psStats = aProj(psObj).psWStats
    If Not (psStats <> -1) Then ASSERT_FAILURE "psStats != nullptr", "src/projectile.cpp:1354", "proj_PostImpactFunc", "Invalid weapon stats pointer": Exit Sub

    age = FromU32(U32(gameTime) - U32(aProj(psObj).born))

    ' Time to finish postimpact effect? (age is compared as unsigned with the unsigned stats)
    If ULT(asWeaponStats(psStats).radiusLife, age) And ULT(asWeaponStats(psStats).upgrade(aProj(psObj).player).periodicalDamageTime, age) Then
        aProj(psObj).state = PROJ_INACTIVE
        Exit Sub
    End If

    ' Periodical damage effect
    If asWeaponStats(psStats).upgrade(aProj(psObj).player).periodicalDamageTime <> 0 Then  ' (unsigned > 0)
        ' See if anything is in the fire and damage it periodically
        proj_checkPeriodicalDamage psObj
    End If
End Sub

' PROJECTILE::update(): the spawned projectile, 0 = none
Public Function PROJECTILE_update(ByVal this As Long) As Long
    Dim psObj As Long, spawnedProjectile As Long, i As Long, n As Long, st As Long

    psObj = this

    checkProjectile psObj, "src/projectile.cpp:1379", "update", max_check_object_recursion

    syncDebugProjectile_ "update", psObj, Asc("<")

    aProj(psObj).prevSpacetime = getSpacetime(psObj)

    ' See if any of the stored objects have died
    ' since the projectile was created
    If aProj(psObj).psSource <> 0 Then
        If SIMPLE_OBJECT_died(aProj(psObj).psSource) <> 0 Then
            syncDebugObject_ "update", aProj(psObj).psSource, Asc("-")
            setProjectileSource psObj, 0
        End If
    End If
    If aProj(psObj).psDest <> 0 Then
        If SIMPLE_OBJECT_died(aProj(psObj).psDest) <> 0 Then
            syncDebugObject_ "update", aProj(psObj).psDest, Asc("-")
            setProjectileDestination psObj, 0
        End If
    End If
    ' Remove dead objects from psDamaged.
    n = 0
    For i = 0 To aProj(psObj).psDamaged_size - 1
        If Not isDead(aProj(psObj).psDamaged(i)) Then
            aProj(psObj).psDamaged(n) = aProj(psObj).psDamaged(i)
            n = n + 1
        End If
    Next
    aProj(psObj).psDamaged_size = n

    ' This extra check fixes a crash in cam2, mission1
    If worldOnMap(aProj(psObj).pos.x, aProj(psObj).pos.y) = False Then
        aProj(psObj).died = 1
        PROJECTILE_update = 0
        Exit Function
    End If

    spawnedProjectile = 0
    st = aProj(psObj).state
    If st = PROJ_INFLIGHT Then
        spawnedProjectile = proj_InFlightFunc(psObj)
        If aProj(psObj).state <> PROJ_IMPACT Then
            GoTo switchEnd
        End If
        ' fallthrough
        st = PROJ_IMPACT
    End If
    If st = PROJ_IMPACT Then
        proj_ImpactFunc psObj
        If aProj(psObj).state <> PROJ_POSTIMPACT Then
            GoTo switchEnd
        End If
        ' fallthrough
        st = PROJ_POSTIMPACT
    End If
    If st = PROJ_POSTIMPACT Then
        proj_PostImpactFunc psObj
    ElseIf st = PROJ_INACTIVE Then
        aProj(psObj).died = aProj(psObj).time_
    End If
switchEnd:

    syncDebugProjectile_ "update", psObj, Asc(">")

    PROJECTILE_update = spawnedProjectile
End Function

' iterate through all projectiles and update their status
Public Sub proj_UpdateAll()
    Dim i As Long, nSpawned As Long, spawned As Long, n As Long, p As Long, keep As Boolean

    ' WZ_PROFILE_SCOPE(proj_UpdateAll)

    nSpawned = 0

    ' Update all projectiles.
    ' Penetrating projectiles may spawn additional projectiles,
    ' which will be returned from `PROJECTILE::update()`.
    ' These need to be added separately to `psProjectileList` later.
    For i = 0 To psProjectileList_size - 1
        spawned = PROJECTILE_update(psProjectileList(i))
        If spawned <> 0 Then
            If nSpawned = 0 Then
                ReDim spawnedProjectiles(0 To 15)
            ElseIf nSpawned > UBound(spawnedProjectiles) Then
                ReDim Preserve spawnedProjectiles(0 To 2 * nSpawned - 1)
            End If
            spawnedProjectiles(nSpawned) = spawned
            nSpawned = nSpawned + 1
        End If
    Next

    ' Remove and free dead projectiles.
    n = 0
    For i = 0 To psProjectileList_size - 1
        p = psProjectileList(i)
        keep = aProj(p).died = 0
        If Not keep Then keep = Not ULT(aProj(p).died, FromU32(U32(gameTime) - U32(deltaGameTime)))
        If keep Then
            psProjectileList(n) = p
            n = n + 1
        Else
            If Not (ProjectileContainer_find(p)) Then ASSERT_FAILURE "it != globalProjectileStorage.end()", "src/projectile.cpp:1471", "proj_UpdateAll", "Invalid projectile, not found in global storage"

            ' Make sure to get rid of some final references in the sound code to this object first
            audio_RemoveObj p

            ProjectileContainer_erase p
        End If
    Next
    psProjectileList_size = n

    ' Add spawned penetrating projectiles,
    ' which were collected earlier during the update procedure.
    For i = 0 To nSpawned - 1
        proj_AddActiveProjectile spawnedProjectiles(i)
    Next
End Sub

Private Sub proj_checkPeriodicalDamage(ByVal psProj As Long)
    Dim psStats As Long, n As Long, gi As Long, psCurr As Long, damageRate As Long, sDamage As DAMAGE_, player As Long

    checkProjectile psProj, "src/projectile.cpp:1490", "proj_checkPeriodicalDamage", max_check_object_recursion

    psStats = aProj(psProj).psWStats
    player = aProj(psProj).player

    n = gridStartIterate(aProj(psProj).pos.x, aProj(psProj).pos.y, asWeaponStats(psStats).upgrade(player).periodicalDamageRadius, gridListPeriodical)
    For gi = 0 To n - 1
        psCurr = gridListPeriodical(gi)
        If SIMPLE_OBJECT_died(psCurr) <> 0 Then
            syncDebugObject_ "proj_checkPeriodicalDamage", psCurr, Asc("-")
            GoTo continue_  ' Do not damage dead objects further.
        End If

        If aiCheckAlliances(player, SIMPLE_OBJECT_player(psCurr)) Then
            GoTo continue_  ' Don't damage your own droids, nor ally droids - unrealistic, but better.
        End If

        If SIMPLE_OBJECT_type(psCurr) = OBJ_DROID Then
            If DROID_isFlying(psCurr) Then
                GoTo continue_  ' Can't set flying units on fire.
            End If
        End If

        If SIMPLE_OBJECT_type(psCurr) = OBJ_FEATURE Then
            If Not asFeatureStats(aFeature(psCurr).psStats).damageable Then
                GoTo continue_  ' Can't destroy oil wells.
            End If
        End If

        If objPeriodicalDamageStart(psCurr) <> gameTime Then
            setObjPeriodicalDamageStart psCurr, gameTime
            BASE_OBJECT_periodicalDamage(psCurr) = 0  ' Reset periodical damage done this tick.
        End If
        damageRate = calcDamage(weaponPeriodicalDamage(asWeaponStats(psStats), player), asWeaponStats(psStats).periodicalDamageWeaponEffect, psCurr)
        debug_ 1526, LOG_NEVER, "proj_checkPeriodicalDamage", "Periodical damage of %d per second to object %d, player %d\n", damageRate, SIMPLE_OBJECT_id(psCurr), SIMPLE_OBJECT_player(psCurr)

        sDamage.psProjectile = psProj
        sDamage.psDest = psCurr
        sDamage.damage = damageRate
        sDamage.weaponClass = asWeaponStats(psStats).periodicalDamageWeaponClass
        sDamage.weaponSubClass = asWeaponStats(psStats).periodicalDamageWeaponSubClass
        sDamage.impactTime = FromU32(U32(gameTime) - Int(U32(deltaGameTime) / 2) + 1)
        sDamage.isDamagePerSecond = True
        sDamage.minDamage = asWeaponStats(psStats).upgrade(player).minimumDamage
        sDamage.empRadiusHit = False

        objectDamage sDamage
continue_:
    Next
End Sub

' return whether a weapon is direct or indirect
Public Function proj_Direct(ByVal psStats As Long) As Boolean
    If Not (psStats <> -1) Then ASSERT_FAILURE "psStats", "src/projectile.cpp:1549", "proj_Direct", "Called with NULL weapon": proj_Direct = False: Exit Function

    Select Case asWeaponStats(psStats).movementModel
    Case MM_DIRECT, MM_HOMINGDIRECT
        proj_Direct = True
        Exit Function
    Case MM_INDIRECT, MM_HOMINGINDIRECT
        proj_Direct = False
        Exit Function
    End Select

    proj_Direct = False ' just to satisfy compiler
End Function

' return the maximum range for a weapon
Public Function proj_GetLongRange(psStats As WEAPON_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/projectile.cpp:1572", "proj_GetLongRange", "Invalid player: %u", player: proj_GetLongRange = 0: Exit Function
    proj_GetLongRange = psStats.upgrade(player).maxRange
End Function

' return the minimum range for a weapon
Public Function proj_GetMinRange(psStats As WEAPON_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/projectile.cpp:1579", "proj_GetMinRange", "Invalid player: %u", player: proj_GetMinRange = 0: Exit Function
    proj_GetMinRange = psStats.upgrade(player).minRange
End Function

' return the short range for a weapon
Public Function proj_GetShortRange(psStats As WEAPON_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/projectile.cpp:1586", "proj_GetShortRange", "Invalid player: %u", player: proj_GetShortRange = 0: Exit Function
    proj_GetShortRange = psStats.upgrade(player).shortRange
End Function

Public Function establishTargetShape(ByVal psTarget As Long) As ObjectShape
    Dim size As Vector2i

    checkObject psTarget, "src/projectile.cpp:1593", "establishTargetShape", max_check_object_recursion

    Select Case SIMPLE_OBJECT_type(psTarget)
    Case OBJ_DROID  ' Circular.
        Select Case aDroid(castDroid(psTarget)).droidType
        Case DROID_WEAPON, DROID_SENSOR, DROID_ECM, DROID_CONSTRUCT, DROID_COMMAND, DROID_REPAIR, DROID_PERSON, _
             DROID_CYBORG, DROID_CYBORG_CONSTRUCT, DROID_CYBORG_REPAIR, DROID_CYBORG_SUPER
            'Watermelon:'hitbox' size is now based on imd size
            establishTargetShape = ObjectShape_radius(MUL32(Abs(iIMDBaseShape_radius(BASE_OBJECT_sDisplay_imd(psTarget))), 2))
            Exit Function
        Case Else   ' DROID_DEFAULT, DROID_TRANSPORTER, DROID_SUPERTRANSPORTER
            establishTargetShape = ObjectShape_radius(TILE_UNITS \ 4) ' how will we arrive at this?
            Exit Function
        End Select
    Case OBJ_STRUCTURE  ' Rectangular.
        size = STRUCTURE_size(castStructure(psTarget))
        establishTargetShape = ObjectShape_size(Vector2i_(size.x * TILE_UNITS \ 2, size.y * TILE_UNITS \ 2))
        Exit Function
    Case OBJ_FEATURE  ' Rectangular.
        establishTargetShape = ObjectShape_size(Vector2i_(asFeatureStats(aFeature(castFeature(psTarget)).psStats).baseWidth * TILE_UNITS \ 2, _
                                                          asFeatureStats(aFeature(castFeature(psTarget)).psStats).baseBreadth * TILE_UNITS \ 2))
        Exit Function
    Case OBJ_PROJECTILE  ' Circular, but can't happen since a PROJECTILE isn't a BASE_OBJECT.
        'Watermelon 1/2 radius of a droid?
        establishTargetShape = ObjectShape_radius(TILE_UNITS \ 8)
        Exit Function
    Case Else
    End Select

    establishTargetShape = ObjectShape_radius(0)  ' Huh?
End Function

' the damage depends on the weapon effect and the target propulsion type or
' structure strength (UDWORD arithmetic, wrapping)
Public Function calcDamage(ByVal baseDamage As Long, ByVal weaponEffect As WEAPON_EFFECT, ByVal psTarget As Long) As Long
    Dim damage As Long, propulsion As Long, body As Long, q As Long

    If baseDamage = 0 Then
        calcDamage = 0
        Exit Function
    End If

    damage = MUL32(baseDamage, 100)

    If SIMPLE_OBJECT_type(psTarget) = OBJ_STRUCTURE Then
        damage = FromU32(U32(damage) + U32(MUL32(baseDamage, asStructStrengthModifier(weaponEffect, asStructureStats(aStruct(psTarget).pStructureType).strength) - 100)))
    ElseIf SIMPLE_OBJECT_type(psTarget) = OBJ_DROID Then
        propulsion = asPropulsionStats(DROID_getPropulsionStats(psTarget)).propulsionType
        body = asBodyStats(DROID_getBodyStats(psTarget)).size
        damage = FromU32(U32(damage) + U32(MUL32(baseDamage, asWeaponModifier(weaponEffect, propulsion) - 100)))
        damage = FromU32(U32(damage) + U32(MUL32(baseDamage, asWeaponModifierBody(weaponEffect, body) - 100)))
    End If

    'Always do at least one damage.
    q = FromU32(Int(U32(damage) / 100))
    If ULT(q, 1) Then q = 1
    calcDamage = q
End Function

' A quick explanation about hown this function works:
'  - It returns an integer between 0 and 100 (see note for exceptions);
'  - this represents the amount of damage inflicted on the droid by the weapon
'    in relation to its original health.
'  - e.g. If 100 points of (*actual*) damage were done to a unit who started
'    off (when first produced) with 400 points then .25 would be returned.
'  - If the actual damage done to a unit is greater than its remaining points
'    then the actual damage is clipped: so if we did 200 actual points of
'    damage to a cyborg with 150 points left the actual damage would be taken
'    as 150.
'  - Should sufficient damage be done to destroy/kill a unit then the value is
'    multiplied by -1, resulting in a negative number. Killed features do not
'    result in negative numbers.
Private Function objectDamageDispatch(psDamage As DAMAGE_) As Long
    Select Case SIMPLE_OBJECT_type(psDamage.psDest)
    Case OBJ_DROID
        objectDamageDispatch = droidDamage(psDamage.psDest, psDamage.psProjectile, psDamage.damage, psDamage.weaponClass, psDamage.weaponSubClass, psDamage.impactTime, psDamage.isDamagePerSecond, psDamage.minDamage, psDamage.empRadiusHit)
        Exit Function

    Case OBJ_STRUCTURE
        objectDamageDispatch = structureDamage(psDamage.psDest, psDamage.psProjectile, psDamage.damage, psDamage.weaponClass, psDamage.weaponSubClass, psDamage.impactTime, psDamage.isDamagePerSecond, psDamage.minDamage, psDamage.empRadiusHit)
        Exit Function

    Case OBJ_FEATURE
        objectDamageDispatch = featureDamage(psDamage.psDest, psDamage.damage, psDamage.weaponClass, psDamage.weaponSubClass, psDamage.impactTime, psDamage.isDamagePerSecond, psDamage.minDamage, psDamage.empRadiusHit)
        Exit Function

    Case OBJ_PROJECTILE
        ASSERT_FAILURE "!""invalid object type: bullet""", "src/projectile.cpp:1694", "objectDamageDispatch", "invalid object type: OBJ_PROJECTILE (id=%d)", SIMPLE_OBJECT_id(psDamage.psDest)

    Case Else
        ASSERT_FAILURE "!""unknown object type""", "src/projectile.cpp:1698", "objectDamageDispatch", "unknown object type %d, id=%d", SIMPLE_OBJECT_type(psDamage.psDest), SIMPLE_OBJECT_id(psDamage.psDest)
    End Select
    objectDamageDispatch = 0
End Function

Private Function isFriendlyFire(psDamage As DAMAGE_) As Boolean
    If psDamage.psDest <> 0 Then
        isFriendlyFire = SIMPLE_OBJECT_player(aProj(psDamage.psProjectile).psSource) = SIMPLE_OBJECT_player(psDamage.psDest)
    End If
End Function

Private Function shouldIncreaseExperience(psDamage As DAMAGE_) As Boolean
    If aProj(psDamage.psProjectile).psSource <> 0 Then
        If Not isFeature(aProj(psDamage.psProjectile).psDest) Then
            shouldIncreaseExperience = Not isFriendlyFire(psDamage)
        End If
    End If
End Function

Private Sub updateKills(psDamage As DAMAGE_)
    Dim psSource As Long, psDroid As Long, psCommander As Long

    psSource = aProj(psDamage.psProjectile).psSource
    If bMultiPlayer Then
        updateMultiStatsKills psDamage.psDest, SIMPLE_OBJECT_player(psSource)
    End If

    If SIMPLE_OBJECT_type(psSource) = OBJ_DROID Then
        psDroid = psSource

        aDroid(psDroid).kills = FromU32(U32(aDroid(psDroid).kills) + 1)

        If hasCommander(psDroid) Then
            psCommander = aGroup(aDroid(psDroid).psGroup).psCommander
            aDroid(psCommander).kills = FromU32(U32(aDroid(psCommander).kills) + 1)
        End If
    ElseIf SIMPLE_OBJECT_type(psSource) = OBJ_STRUCTURE Then
        psCommander = getDesignatorAttackingObject(SIMPLE_OBJECT_player(psSource), aProj(psDamage.psProjectile).psDest)

        If psCommander <> 0 Then
            aDroid(psCommander).kills = FromU32(U32(aDroid(psCommander).kills) + 1)
        End If
    End If
End Sub

Private Function objectDamage(psDamage As DAMAGE_) As Long
    Dim relativeDamage As Long, isTargetDestroyed As Boolean

    relativeDamage = objectDamageDispatch(psDamage)

    If shouldIncreaseExperience(psDamage) Then
        proj_UpdateExperience psDamage.psProjectile, MUL32(Abs(relativeDamage), getExpGain(SIMPLE_OBJECT_player(aProj(psDamage.psProjectile).psSource))) \ 100

        isTargetDestroyed = relativeDamage < 0

        If isTargetDestroyed Then
            updateKills psDamage
        End If
    End If

    objectDamage = relativeDamage
End Function

' (VB) imd->max.y and imd->min.y of an iIMDBaseShape *
Private Function imdHeightMax(ByVal imd As Long) As Long
    Dim x As Long, y As Long, z As Long
    iIMDBaseShape_max imd, x, y, z
    imdHeightMax = y
End Function

Private Function imdHeightMin(ByVal imd As Long) As Long
    Dim x As Long, y As Long, z As Long
    iIMDBaseShape_min imd, x, y, z
    imdHeightMin = y
End Function

Public Function establishTargetHeight(ByVal psTarget As Long) As Long
    Dim psDroid As Long, height As Long, utilityHeight As Double, yMax As Long, yMin As Long, imd As Long
    Dim psStructureStats As Long

    If psTarget = 0 Then
        establishTargetHeight = 0
        Exit Function
    End If

    checkObject psTarget, "src/projectile.cpp:1833", "establishTargetHeight", max_check_object_recursion

    Select Case SIMPLE_OBJECT_type(psTarget)
    Case OBJ_DROID
        psDroid = psTarget
        imd = asBodyStats(DROID_getBodyStats(psDroid)).pIMD
        height = imdHeightMax(imd) - imdHeightMin(imd)                ' unsigned
        yMax = 0: yMin = 0  ' Temporaries for addition of utility's height to total height

        ' VTOL's don't have pIMD either it seems...
        If DROID_isVtol(psDroid) Then
            establishTargetHeight = FromU32(U32(height) + VTOL_HITBOX_MODIFICATOR)
            Exit Function
        End If

        Select Case aDroid(psDroid).droidType
        Case DROID_WEAPON
            If aDroid(psDroid).numWeaps > 0 Then
                imd = asWeaponStats(DROID_getWeaponStats(psDroid, 0)).pIMD
                ' Don't do this for Barbarian Propulsions as they don't possess a turret (and thus have pIMD == NULL)
                If imd = 0 Then
                    establishTargetHeight = height
                    Exit Function
                End If

                yMax = imdHeightMax(imd)
                yMin = imdHeightMin(imd)
            End If

        Case DROID_SENSOR
            imd = asSensorStats(DROID_getSensorStats(psDroid)).pIMD
            yMax = imdHeightMax(imd)
            yMin = imdHeightMin(imd)
        Case DROID_ECM
            imd = asECMStats(DROID_getECMStats(psDroid)).pIMD
            yMax = imdHeightMax(imd)
            yMin = imdHeightMin(imd)
        Case DROID_CONSTRUCT
            imd = asConstructStats(DROID_getConstructStats(psDroid)).pIMD
            yMax = imdHeightMax(imd)
            yMin = imdHeightMin(imd)
        Case DROID_REPAIR
            imd = asRepairStats(DROID_getRepairStats(psDroid)).pIMD
            yMax = imdHeightMax(imd)
            yMin = imdHeightMin(imd)
        Case DROID_PERSON, DROID_CYBORG, DROID_CYBORG_CONSTRUCT, DROID_CYBORG_REPAIR, DROID_CYBORG_SUPER, _
             DROID_DEFAULT, DROID_TRANSPORTER, DROID_SUPERTRANSPORTER, DROID_COMMAND, DROID_ANY
            ' TODO:add person 'state'checks here(stand, knee, crouch, prone etc)
            ' Commanders don't have pIMD either
            establishTargetHeight = height
            Exit Function
        End Select

        ' TODO: check the /2 - does this really make sense? why + ?
        utilityHeight = Int(U32(FromU32(CDbl(yMax) + yMin)) / 2)    ' unsigned

        establishTargetHeight = FromU32(U32(height) + utilityHeight)
    Case OBJ_STRUCTURE
        psStructureStats = aStruct(psTarget).pStructureType
        height = imdHeightMax(asStructureStats(psStructureStats).pIMD(0)) + imdHeightMin(asStructureStats(psStructureStats).pIMD(0))
        height = height - gateCurrentOpenHeight(psTarget, gameTime, 2)  ' Treat gate as at least 2 units tall, even if open, so that it's possible to hit.
        establishTargetHeight = height
    Case OBJ_FEATURE
        ' Just use imd ymax+ymin
        establishTargetHeight = imdHeightMax(aFeature(psTarget).sDisplay.imd) + imdHeightMin(aFeature(psTarget).sDisplay.imd)
    Case OBJ_PROJECTILE
        establishTargetHeight = BULLET_FLIGHT_HEIGHT
    Case Else
        establishTargetHeight = 0
    End Select
End Function

' CHECK_PROJECTILE(object) is checkProjectile object, "<file>:<line>", "<function>", max_check_object_recursion.
' (VB) as checkDroid, its ASSERT_HELPERs log under the caller's location
Public Sub checkProjectile(ByVal psProjectile As Long, ByVal location_description As String, ByVal function_ As String, ByVal recurse As Long)
    Dim n As Long, s As Long

    If recurse < 0 Then
        Exit Sub
    End If

    If Not (psProjectile <> 0) Then ASSERT_HELPER_FAILURE 1938, "psProjectile != nullptr", location_description, function_, "CHECK_PROJECTILE: NULL pointer": Exit Sub
    If Not (aProj(psProjectile).psWStats <> -1) Then ASSERT_HELPER_FAILURE 1939, "psProjectile->psWStats != nullptr", location_description, function_, "CHECK_PROJECTILE"
    If Not (aProj(psProjectile).typ = OBJ_PROJECTILE) Then ASSERT_HELPER_FAILURE 1940, "psProjectile->type == OBJ_PROJECTILE", location_description, function_, "CHECK_PROJECTILE"
    If Not (aProj(psProjectile).player < MAX_PLAYERS) Then ASSERT_HELPER_FAILURE 1941, "psProjectile->player < MAX_PLAYERS", location_description, function_, "CHECK_PROJECTILE: Out of bound owning player number (%u)", aProj(psProjectile).player
    s = aProj(psProjectile).state
    If Not (s = PROJ_INFLIGHT Or s = PROJ_IMPACT Or s = PROJ_POSTIMPACT Or s = PROJ_INACTIVE) Then ASSERT_HELPER_FAILURE 1942, "psProjectile->state == PROJ_INFLIGHT || psProjectile->state == PROJ_IMPACT || psProjectile->state == PROJ_POSTIMPACT || psProjectile->state == PROJ_INACTIVE", location_description, function_, "CHECK_PROJECTILE: invalid projectile state: %u", s

    If aProj(psProjectile).psDest <> 0 Then
        checkObject aProj(psProjectile).psDest, location_description, function_, recurse - 1
    End If

    If aProj(psProjectile).psSource <> 0 Then
        checkObject aProj(psProjectile).psSource, location_description, function_, recurse - 1
    End If

    For n = 0 To aProj(psProjectile).psDamaged_size - 1
        checkObject aProj(psProjectile).psDamaged(n), location_description, function_, recurse - 1
    Next
End Sub
