Attribute VB_Name = "effects_c"
' src/effects.cpp, effects.h: the spot effects (explosions, smoke, gravitons,
' construction dust, blood, destruction, fire, the satellite laser,
' fireworks, waypoints, droid death animations), whole.
' The container (static PagedEntityContainer<EFFECT> gActiveEffects) is the
' pool aEffect with its alive flags: an EFFECT * is a 1-based handle into it
' (0 = nullptr), as bucket3d's RENDER_EFFECT entries and renderEffect take
' it; emplace takes a freed slot again, the earliest freed first (the C++
' page's free index queue; one queue here, as for one page), else the next
' new one; iteration runs over the slots in order. An iIMDShape * (the
' display model) is WzGfx's handle (0 = nullptr), read through GfxApi's
' iIMDShape_* accessors. A LIGHT goes to WzGfx's current LightingData
' (LightingData_lights_push_back): processEffects' lightData is
' getCurrentLightingData(), the only one drawTiles passes. The overload
' addEffect(pos, group, type, specified, imd, lit) is addEffect with
' graphicsTime (its body); the one with rot and velocity is addEffect_velocity
' (velocity only) and addEffect_rotVelocity. The display randomness is
' cstdlib_vb's rand, as C++'s rand(). The uint8_t / uint16_t fields are Longs
' masked where C++ would truncate.
' (VB) vbEffectsOff: tools/run_tests.py's vbtest runs set it, where statsref's
' stand-ins add no effects (and the misc models are not loaded): addEffect and
' the other entry points the simulation calls then do nothing.
Option Explicit

Public Const SHOCK_WAVE_HEIGHT As Long = 64

' All the effect groups
Public Enum EFFECT_GROUP
    EFFECT_EXPLOSION
    EFFECT_CONSTRUCTION
    EFFECT_SMOKE
    EFFECT_GRAVITON
    EFFECT_WAYPOINT
    EFFECT_BLOOD
    EFFECT_DESTRUCTION
    EFFECT_SAT_LASER
    EFFECT_FIRE
    EFFECT_FIREWORK
    EFFECT_DROID_ANIMEVENT_DYING
    EFFECT_FREED
End Enum

' Might not even need this
Public Enum EFFECT_TYPE
    EXPLOSION_TYPE_SMALL
    EXPLOSION_TYPE_VERY_SMALL
    EXPLOSION_TYPE_MEDIUM
    EXPLOSION_TYPE_LARGE
    EXPLOSION_TYPE_SPECIFIED                ' superfluous?
    EXPLOSION_TYPE_NOT_FACING
    EXPLOSION_TYPE_SPECIFIED_SOLID
    EXPLOSION_TYPE_SPECIFIED_FIXME
    EXPLOSION_TYPE_FLAMETHROWER
    EXPLOSION_TYPE_LASER
    EXPLOSION_TYPE_TESLA
    EXPLOSION_TYPE_DISCOVERY
    EXPLOSION_TYPE_FLARE
    EXPLOSION_TYPE_PLASMA
    EXPLOSION_TYPE_KICKUP
    EXPLOSION_TYPE_LAND_LIGHT
    EXPLOSION_TYPE_SHOCKWAVE
    'EXPLOSION_TYPE_SKYSCRAPER,
    GRAVITON_TYPE_STANDARD
    GRAVITON_TYPE_EMITTING_DR
    GRAVITON_TYPE_EMITTING_ST
    GRAVITON_TYPE_GIBLET
    SMOKE_TYPE_DRIFTING
    SMOKE_TYPE_DRIFTING_HIGH
    SMOKE_TYPE_DRIFTING_SMALL
    SMOKE_TYPE_BILLOW
    SMOKE_TYPE_STEAM
    SMOKE_TYPE_TRAIL
    FIRE_TYPE_LOCALISED
    FIRE_TYPE_SMOKY
    FIRE_TYPE_SMOKY_BLUE
    CONSTRUCTION_TYPE_DRIFTING
    BLOOD_TYPE_NORMAL
    DUST_TYPE_NORMAL
    DESTRUCTION_TYPE_DROID
    DESTRUCTION_TYPE_STRUCTURE
    DESTRUCTION_TYPE_POWER_STATION
    DESTRUCTION_TYPE_WALL_SECTION
    DESTRUCTION_TYPE_FEATURE
    DESTRUCTION_TYPE_SKYSCRAPER
    SAT_LASER_STANDARD
    WAYPOINT_TYPE
    FIREWORK_TYPE_STARBURST
    FIREWORK_TYPE_LAUNCHER
    DROID_ANIMEVENT_DYING_NORMAL
    DROID_ANIMEVENT_DYING_NORMAL_ST
End Enum

Public Enum LAND_LIGHT_SPEC
    LL_MIDDLE
    LL_INNER
    LL_OUTER
End Enum

Public Type EFFECT
    position As Vector3f                ' = { 0.f, 0.f, 0.f }: world coordinates of the effect - floats on the PC.
    velocity As Vector3f                ' movement values per update
    rotation As Vector3i                ' current rotation - only for gravitons
    spin As Vector3i                    ' rotation info for spinning things.
    imd As Long                         ' const iIMDShape *: pointer to the imd the effect uses (= nullptr)
    birthTime As Long                   ' uint32_t: what time was it introduced into the world?
    lastFrame As Long                   ' uint32_t: when did we last update the frame?
    group As EFFECT_GROUP               ' = EFFECT_FREED: what group is it - explosion, building effect etc....
    typ As EFFECT_TYPE                  ' = EXPLOSION_TYPE_SMALL: what type is it within the group?
    size As Long                        ' uint16_t: Size in terms of percent of original imd.
    frameDelay As Long                  ' uint16_t: how many game ticks between each frame?
    lifeSpan As Long                    ' uint16_t: what is it's life expectancy?
    radius As Long                      ' uint16_t: Used for area effects
    player As Long                      ' uint8_t = MAX_PLAYERS: when the effect in question needs a player's color
    control As Long                     ' uint8_t: Controls the bits above - essential,flips etc
    frameNumber As Long                 ' uint8_t: what frame number is the imd on?
    baseScale As Long                   ' uint8_t: if scaled, what's bottom line?
    specific As Long                    ' uint8_t: how many times has it bounced?
End Type

Private Const GRAVITON_GRAVITY As Single = -800!
Private Const EFFECT_X_FLIP As Long = &H1
Private Const EFFECT_Y_FLIP As Long = &H2
Private Const EFFECT_CYCLIC As Long = &H4
Private Const EFFECT_ESSENTIAL As Long = &H8
Private Const EFFECT_FACING As Long = &H10
Private Const EFFECT_SCALED As Long = &H20
Private Const EFFECT_LIT As Long = &H40

Private Const TRAIL_SMOKE_LIFESPAN As Long = 1200
Private Const CONSTRUCTION_LIFESPAN As Long = 5000

Private Const EXPLOSION_TESLA_FRAME_DELAY As Long = 65
Private Const EXPLOSION_PLASMA_FRAME_DELAY As Long = 45
Private Const BLOOD_FRAME_DELAY As Long = 150
Private Const DESTRUCTION_FRAME_DELAY As Long = 200

Private Const TESLA_SPEED As Long = 170 ' + (30 - rand()%60))
Private Const TESLA_SIZE As Long = 100  ' + (20 - rand()%40))

Private Const GIBLET_INIT_VEL_Y As Single = 12!

Private Const DROID_DESTRUCTION_DURATION As Long = 3 * GAME_TICKS_PER_SEC \ 2          ' 1.5 seconds
Private Const STRUCTURE_DESTRUCTION_DURATION As Long = (7 * GAME_TICKS_PER_SEC) \ 2   ' 3.5 seconds

Private Const EFFECT_EXPLOSION_ADDITIVE As Long = 164
Private Const EFFECT_PLASMA_ADDITIVE As Long = 224
Private Const EFFECT_SMOKE_TRANSPARENCY As Long = 130
Private Const EFFECT_STEAM_TRANSPARENCY As Long = 128
Private Const EFFECT_BLOOD_TRANSPARENCY As Long = 128

Private Const EFFECT_DROID_DIVISION As Long = 101
Private Const EFFECT_STRUCTURE_DIVISION As Long = 103

Private Const DROID_UPDATE_INTERVAL As Long = 500
Private Const STRUCTURE_UPDATE_INTERVAL As Long = 1250
Private Const BASE_FLAME_SIZE As Long = 80
Private Const BASE_LASER_SIZE As Long = 10
Private Const BASE_PLASMA_SIZE As Long = 0
Private Const DISCOVERY_SIZE As Long = 60
Private Const FLARE_SIZE As Long = 100
Private Const SHOCKWAVE_SPEED As Long = GAME_TICKS_PER_SEC
Private Const MAX_SHOCKWAVE_SIZE As Long = 500

' lib/ivis_opengl/pietypes.h's render style flags, imd.h's iV_IMD_PREMULTIPLIED
Private Const pie_TRANSLUCENT As Long = &H2
Private Const pie_ADDITIVE As Long = &H4
Private Const pie_RAISE As Long = &H20
Private Const pie_SHADOW As Long = &H80
Private Const pie_PREMULTIPLIED As Long = &H200
Private Const pie_NODEPTHWRITE As Long = &H400
Private Const pie_RAISE_SCALE As Long = 256
Private Const iV_IMD_PREMULTIPLIED As Long = &H4

' display3ddef.h: SKY_SHIMMY_BASE ((DEG(1)*SKY_MULT)/2), SKY_MULT 1
Private Const SKY_SHIMMY_BASE As Long = 182

' bucket3d.h's RENDER_EFFECT (the enum is fork S's bucket3d_c)
' static PagedEntityContainer<EFFECT> gActiveEffects;
Public aEffect() As EFFECT              ' (VB) the container's storage, 1-based
Private effectAlive() As Boolean
Private effectCap As Long
Private effectUsed As Long              ' slots taken since the last clear() (the highest index)
Private effectSize As Long              ' gActiveEffects.size()
Private freeQueue() As Long             ' the recycled slots, earliest freed first (a circular queue)
Private freeHead As Long
Private freeCount As Long
Private effectEmpty As EFFECT

' Tick counts for updates on a particular interval
Private lastUpdateStructures(0 To EFFECT_STRUCTURE_DIVISION - 1) As Long  ' UDWORD

Private auxVar As Long                  ' UDWORD: dirty filthy hack - don't look for what this does.... //FIXME
Private auxVarSec As Long               ' UDWORD: dirty filthy hack - don't look for what this does.... //FIXME
Private specifiedSize As Long           ' UDWORD
Private ellSpec As Long                 ' UDWORD
Private EffectForPlayer As Long         ' uint8_t (= 0)

Public vbEffectsOff As Boolean          ' (VB) see the module comment

'--------------------------------------------------------------------------
' (VB) the container, the macros and the casts

' EFFECT's default member initialisers
Private Function EFFECT_new() As EFFECT
    Dim e As EFFECT
    e = effectEmpty
    e.group = EFFECT_FREED
    e.typ = EXPLOSION_TYPE_SMALL
    e.player = MAX_PLAYERS
    EFFECT_new = e
End Function

' gActiveEffects.emplace(std::move(effect)): its handle
Private Function gActiveEffects_emplace(e As EFFECT) As Long
    Dim h As Long, n As Long
    If freeCount > 0 Then
        h = freeQueue(freeHead)
        freeHead = (freeHead + 1) Mod effectCap
        freeCount = freeCount - 1
    Else
        If effectUsed = effectCap Then
            n = effectCap
            If effectCap = 0 Then effectCap = 1024 Else effectCap = effectCap * 2
            ReDim Preserve aEffect(1 To effectCap)
            ReDim Preserve effectAlive(1 To effectCap)
            growFreeQueue n
        End If
        effectUsed = effectUsed + 1
        h = effectUsed
    End If
    aEffect(h) = e
    effectAlive(h) = True
    effectSize = effectSize + 1
    gActiveEffects_emplace = h
End Function

' the free queue keeps its order when the capacity grows
Private Sub growFreeQueue(ByVal oldCap As Long)
    Dim q() As Long, i As Long
    If freeCount > 0 Then
        ReDim q(0 To freeCount - 1)
        For i = 0 To freeCount - 1
            q(i) = freeQueue((freeHead + i) Mod oldCap)
        Next
    End If
    ReDim freeQueue(0 To effectCap - 1)
    For i = 0 To freeCount - 1
        freeQueue(i) = q(i)
    Next
    freeHead = 0
End Sub

' gActiveEffects.erase(it)
Private Sub gActiveEffects_erase(ByVal h As Long)
    If Not effectAlive(h) Then Exit Sub
    effectAlive(h) = False
    effectSize = effectSize - 1
    freeQueue((freeHead + freeCount) Mod effectCap) = h
    freeCount = freeCount + 1
End Sub

' gActiveEffects.clear()
Private Sub gActiveEffects_clear()
    Erase aEffect
    Erase effectAlive
    Erase freeQueue
    effectCap = 0
    effectUsed = 0
    effectSize = 0
    freeHead = 0
    freeCount = 0
End Sub

' (VB) for bucket3d and the tests: whether a handle is a live effect
Public Function EFFECT_isAlive(ByVal h As Long) As Boolean
    If h < 1 Or h > effectUsed Then Exit Function
    EFFECT_isAlive = effectAlive(h)
End Function

' (static_cast<int>(float): math_vb's ToInt, which truncates and gives INT_MIN out of range, as x86)

' PERCENT(a,b) (((a)*100)/(b)) on UDWORDs
Private Function PERCENT_u(ByVal a As Long, ByVal b As Long) As Long
    Dim p As Double
    p = U32(a) * 100#
    p = p - Int(p / 4294967296#) * 4294967296#
    PERCENT_u = FromU32(Int(p / U32(b)))
End Function

Private Function ONEINTWO() As Boolean
    ONEINTWO = ((rand() Mod 2) = 0)
End Function

Private Function SKY_SHIMMY() As Long
    SKY_SHIMMY = SKY_SHIMMY_BASE - (rand() Mod (2 * SKY_SHIMMY_BASE))
End Function

Private Function NORMAL_SMOKE_LIFESPAN() As Long
    NORMAL_SMOKE_LIFESPAN = 6000 + rand() Mod 3000
End Function

Private Function SMALL_SMOKE_LIFESPAN() As Long
    SMALL_SMOKE_LIFESPAN = 3000 + rand() Mod 3000
End Function

Private Function SMOKE_FRAME_DELAY() As Long
    SMOKE_FRAME_DELAY = 40 + rand() Mod 30
End Function

Private Function EXPLOSION_FRAME_DELAY() As Long
    EXPLOSION_FRAME_DELAY = 25 + rand() Mod 40
End Function

Private Function GRAVITON_FRAME_DELAY() As Long
    GRAVITON_FRAME_DELAY = 100 + rand() Mod 50
End Function

Private Function GRAVITON_BLOOD_DELAY() As Long
    GRAVITON_BLOOD_DELAY = 200 + rand() Mod 100
End Function

Private Function CONSTRUCTION_FRAME_DELAY() As Long
    CONSTRUCTION_FRAME_DELAY = 40 + rand() Mod 30
End Function

Private Function EXPLOSION_SIZE() As Long
    EXPLOSION_SIZE = 110 + (30 - rand() Mod 60)
End Function

Private Function BLOOD_SIZE() As Long
    BLOOD_SIZE = 100 + (30 - rand() Mod 60)
End Function

Private Function BLOOD_FALL_SPEED() As Long
    BLOOD_FALL_SPEED = -(20 + rand() Mod 20)
End Function

Private Function GRAVITON_INIT_VEL_X() As Single
    GRAVITON_INIT_VEL_X = CSng(200 - rand() Mod 300)
End Function

Private Function GRAVITON_INIT_VEL_Z() As Single
    GRAVITON_INIT_VEL_Z = CSng(200 - rand() Mod 300)
End Function

Private Function GRAVITON_INIT_VEL_Y() As Single
    GRAVITON_INIT_VEL_Y = CSng(300 + rand() Mod 100)
End Function

Private Function GIBLET_INIT_VEL_X() As Single
    GIBLET_INIT_VEL_X = CSng(50 - rand() Mod 100)
End Function

Private Function GIBLET_INIT_VEL_Z() As Single
    GIBLET_INIT_VEL_Z = CSng(50 - rand() Mod 100)
End Function

' a LIGHT pushed into lightData.lights (its position the Vector3f converted to Vector3i)
Private Sub pushLight(position As Vector3f, ByVal range As Long, ByVal colour As Long)
    LightingData_lights_push_back ToInt(position.x), ToInt(position.y), ToInt(position.z), range, colour
End Sub

'--------------------------------------------------------------------------

Public Sub shutdownEffectsSystem()
    gActiveEffects_clear
End Sub

'!
' * Initialise effects system
Public Sub initEffectsSystem()
    shutdownEffectsSystem
End Sub

Private Function positionEffect(psEffect As EFFECT) As mat4
    ' Establish world position
    Dim dv As Vector3f
    dv = glm_vec3(psEffect.position.x, psEffect.position.y, -(psEffect.position.z))

    positionEffect = glm_translate(dv)
End Function

Public Sub effectSetLandLightSpec(ByVal spec As LAND_LIGHT_SPEC)
    ellSpec = spec
End Sub

Public Sub effectSetSize(ByVal size As Long)
    specifiedSize = size
End Sub

Public Sub addMultiEffect(basePos As Vector3i, scatter As Vector3i, ByVal group As EFFECT_GROUP, ByVal typ As EFFECT_TYPE, ByVal specified As Boolean, ByVal imd As Long, ByVal number As Long, ByVal lit As Boolean, ByVal size As Long, ByVal effectTime As Long)
    Dim i As Long, scatPos As Vector3i
    If vbEffectsOff Then Exit Sub
    If number = 0 Then
        Exit Sub
    End If

    ' Set up the scaling for specified ones
    specifiedSize = size

    ' If there's only one, make sure it's in the centre
    If number = 1 Then
        addEffect basePos, group, typ, specified, imd, -CLng(lit), effectTime
    Else
        ' Fix for jim
        scatter.x = scatter.x \ 10
        scatter.y = scatter.y \ 10
        scatter.z = scatter.z \ 10

        ' There are multiple effects - so scatter them around according to parameter
        For i = 0 To number - 1
            ' This scatters in a cube - is there a good reason for that, or just legacy?
            ' (the Vector3i's arguments taken left to right)
            scatPos.x = rand() Mod (scatter.x * 2 + 1)
            scatPos.y = rand() Mod (scatter.y * 2 + 1)
            scatPos.z = rand() Mod (scatter.z * 2 + 1)
            scatPos.x = basePos.x + scatter.x - scatPos.x
            scatPos.y = basePos.y + scatter.y - scatPos.y
            scatPos.z = basePos.z + scatter.z - scatPos.z
            addEffect scatPos, group, typ, specified, imd, -CLng(lit), effectTime
        Next
    End If
End Sub

' When we need to set the effect for the player's color
Public Sub SetEffectForPlayer(ByVal player As Long)
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < 11", "src/effects.cpp:295", "SetEffectForPlayer", "player is set to a invalid number of %d", player

    EffectForPlayer = getPlayerColour(player) And &HFF&
End Sub

Private Function updateDroidDeathAnimationEffect(psEffect As EFFECT) As Boolean
    Dim imd As Long, deltaX As Single, deltaY As Single, timeAnimationStarted As Long, objanimcycles As Long
    imd = psEffect.imd
    If imd = 0 Then
        updateDroidDeathAnimationEffect = False
        Exit Function
    End If

    ' Move it
    deltaX = graphicsTimeAdjustedIncrement(psEffect.velocity.x)
    deltaY = graphicsTimeAdjustedIncrement(psEffect.velocity.y)
    psEffect.position.x = psEffect.position.x + deltaX
    psEffect.position.z = psEffect.position.z + deltaY

    ' NOTE: We actually store the timeAnimationStarted (in gameTime) in psEffect->lastFrame
    timeAnimationStarted = psEffect.lastFrame
    If iIMDShape_objanimcycles(imd) > 0 Then objanimcycles = iIMDShape_objanimcycles(imd) Else objanimcycles = 1 ' do not allow infinite looping for death animations
    If ULT(timeAnimationStarted + iIMDShape_objanimtime(imd) * objanimcycles, gameTime) Then
        ' Done animating (animation is defined by body - other components should follow suit)
        ' Kill off effect
        updateDroidDeathAnimationEffect = False
        Exit Function
    End If
    updateDroidDeathAnimationEffect = True
End Function

Private Sub renderDroidDeathAnimationEffect(psEffect As EFFECT, viewMatrix As mat4)
    Dim modelMatrix As mat4, brightness As Long, pieFlag As Long, iPieData As Long, strImd As Long, timeAnimationStarted As Long
    modelMatrix = positionEffect(psEffect)
    modelMatrix = glm_mat4_mul(modelMatrix, glm_mat4_mul(glm_mat4_mul( _
        glm_rotate(UNDEG(psEffect.rotation.x), glm_vec3(1!, 0!, 0!)), _
        glm_rotate(UNDEG(psEffect.rotation.y), glm_vec3(0!, 1!, 0!))), _
        glm_rotate(UNDEG(psEffect.rotation.z), glm_vec3(0!, 0!, 1!))))

    ' If original (psDroid->droidType == DROID_PERSON)
    If psEffect.typ = DROID_ANIMEVENT_DYING_NORMAL_ST Then
        ' previously, used to scale the model by 0.75 - no longer needed with the fixed models though...
    End If

    brightness = pal_SetBrightness(255) ' NOTE: brightness is not preserved from the original droid that died - FIXME: ?
    pieFlag = pie_SHADOW
    iPieData = 0

    strImd = psEffect.imd
    timeAnimationStarted = psEffect.lastFrame
    Do While strImd <> 0
        drawShape strImd, timeAnimationStarted, psEffect.player, brightness, pieFlag, iPieData, modelMatrix, viewMatrix
        strImd = iIMDShape_next(strImd)
    Loop
End Sub

' the overload without rot and velocity; effectTime missing: the overload without it, which passes graphicsTime
Public Sub addEffect(pos As Vector3i, ByVal group As EFFECT_GROUP, ByVal typ As EFFECT_TYPE, ByVal specified As Boolean, ByVal imd As Long, ByVal lit As Long, Optional ByVal effectTime As Variant)
    Dim noRot As Vector3i, noVelocity As Vector3f, t As Long
    If IsMissing(effectTime) Then t = graphicsTime Else t = effectTime
    addEffectImpl pos, group, typ, specified, imd, lit, t, False, noRot, False, noVelocity
End Sub

' addEffect(pos, group, type, specified, imd, lit, effectTime, nullptr, &velocity)
Public Sub addEffect_velocity(pos As Vector3i, ByVal group As EFFECT_GROUP, ByVal typ As EFFECT_TYPE, ByVal specified As Boolean, ByVal imd As Long, ByVal lit As Long, ByVal effectTime As Long, velocity As Vector3f)
    Dim noRot As Vector3i
    addEffectImpl pos, group, typ, specified, imd, lit, effectTime, False, noRot, True, velocity
End Sub

' addEffect(pos, group, type, specified, imd, lit, effectTime, &rot, &velocity)
Public Sub addEffect_rotVelocity(pos As Vector3i, ByVal group As EFFECT_GROUP, ByVal typ As EFFECT_TYPE, ByVal specified As Boolean, ByVal imd As Long, ByVal lit As Long, ByVal effectTime As Long, rot As Vector3i, velocity As Vector3f)
    addEffectImpl pos, group, typ, specified, imd, lit, effectTime, True, rot, True, velocity
End Sub

' void addEffect(const Vector3i *pos, EFFECT_GROUP group, EFFECT_TYPE type, bool specified, const iIMDShape *imd, int lit,
'                unsigned effectTime, Vector3i *rot /*= nullptr*/, Vector3f *velocity /*= nullptr*/)
Private Sub addEffectImpl(pos As Vector3i, ByVal group As EFFECT_GROUP, ByVal typ As EFFECT_TYPE, ByVal specified As Boolean, ByVal imd As Long, ByVal lit As Long, ByVal effectTime As Long, ByVal hasRot As Boolean, rot As Vector3i, ByVal hasVelocity As Boolean, velocity As Vector3f)
    Dim e As EFFECT
    If vbEffectsOff Then Exit Sub
    If gamePaused() Then
        Exit Sub
    End If
    e = EFFECT_new()
    ' Reset control bits
    e.control = 0

    ' Store away it's position - into FRACTS
    e.position.x = pos.x
    e.position.y = pos.y
    e.position.z = pos.z

    If hasRot Then
        e.rotation.x = rot.x
        e.rotation.y = rot.y
        e.rotation.z = rot.z
    End If

    If hasVelocity Then
        e.velocity = velocity
    End If

    ' Now, note group and type
    e.group = group
    e.typ = typ

    ' and if the effect needs the player's color for certain things
    e.player = EffectForPlayer
    SetEffectForPlayer 0  ' reset it

    ' Set when it entered the world
    e.birthTime = effectTime
    e.lastFrame = effectTime

    If group = EFFECT_GRAVITON And (typ = GRAVITON_TYPE_GIBLET Or typ = GRAVITON_TYPE_EMITTING_DR) Then
        e.frameNumber = lit And &HFF&
    Else
        ' Starts off on frame zero
        e.frameNumber = 0
    End If

    '
    '   See what kind of effect it is - the add function is different for each,
    '   although some things are shared
    '
    e.imd = 0
    If lit <> 0 Then
        e.control = e.control Or EFFECT_LIT
    End If

    If specified Then
        ' We're specifying what the imd is - override
        e.imd = imd
        e.size = specifiedSize And &HFFFF&
    End If

    ' Do all the effect type specific stuff
    Select Case group
    Case EFFECT_SMOKE
        effectSetupSmoke e
    Case EFFECT_GRAVITON
        effectSetupGraviton e
    Case EFFECT_EXPLOSION
        effectSetupExplosion e
    Case EFFECT_CONSTRUCTION
        effectSetupConstruction e
    Case EFFECT_WAYPOINT
        effectSetupWayPoint e
    Case EFFECT_BLOOD
        effectSetupBlood e
    Case EFFECT_DESTRUCTION
        effectSetupDestruction e
    Case EFFECT_FIRE
        effectSetupFire e
    Case EFFECT_SAT_LASER
        effectSetupSatLaser e
    Case EFFECT_FIREWORK
        effectSetupFirework e
    Case EFFECT_DROID_ANIMEVENT_DYING
        e.lastFrame = gameTime ' store the timeAnimationStarted in lastFrame
    Case EFFECT_FREED
        ASSERT_FAILURE "false", "src/effects.cpp:460", "addEffect", "Weirdy group type for an effect"
    End Select

    ' As of yet, it hasn't bounced (or whatever)...
    If typ <> EXPLOSION_TYPE_LAND_LIGHT Then
        e.specific = 0
    End If

    If Not (e.imd <> 0 Or group = EFFECT_DESTRUCTION Or group = EFFECT_FIRE Or group = EFFECT_SAT_LASER) Then ASSERT_FAILURE "effect.imd != nullptr || group == EFFECT_DESTRUCTION || group == EFFECT_FIRE || group == EFFECT_SAT_LASER", "src/effects.cpp:470", "addEffect", "null effect imd"

    gActiveEffects_emplace e
End Sub


' Calls all the update functions for each different currently active effect
Public Sub processEffects(perspectiveViewMatrix As mat4)
    Dim it As Long, e As EFFECT, keep As Boolean
    it = 1
    Do While it <= effectUsed
        If effectAlive(it) Then
            If Not ULT(graphicsTime, aEffect(it).birthTime) Then   ' Don't process, if it doesn't exist yet
                ' (VB) updated in a copy: the update can add effects, which may grow aEffect (an element
                ' passed ByRef would lock the array: error 10)
                e = aEffect(it)
                keep = updateEffect(e)
                aEffect(it) = e
                If Not keep Then
                    gActiveEffects_erase it
                    GoTo continue_
                End If
                If aEffect(it).group <> EFFECT_FREED Then
                    If clipXY(ToInt(aEffect(it).position.x), ToInt(aEffect(it).position.z)) Then
                        bucketAddTypeToList RENDER_EFFECT, it, perspectiveViewMatrix
                    End If
                End If
            End If
        End If
continue_:
        it = it + 1
    Loop

    ' Add any structure effects
    effectStructureUpdates
End Sub

' The general update function for all effects - calls a specific one for each. Returns false if effect should be deleted.
Private Function updateEffect(psEffect As EFFECT) As Boolean
    ' What type of effect are we dealing with?
    Select Case psEffect.group
    Case EFFECT_EXPLOSION
        updateEffect = updateExplosion(psEffect)
        Exit Function
    Case EFFECT_WAYPOINT
        If Not gamePaused() Then
            updateEffect = updateWaypoint(psEffect)
            Exit Function
        End If
        updateEffect = True
        Exit Function
    Case EFFECT_CONSTRUCTION
        If Not gamePaused() Then
            updateEffect = updateConstruction(psEffect)
            Exit Function
        End If
        updateEffect = True
        Exit Function
    Case EFFECT_SMOKE
        If Not gamePaused() Then
            updateEffect = updatePolySmoke(psEffect)
            Exit Function
        End If
        updateEffect = True
        Exit Function
    Case EFFECT_GRAVITON
        If Not gamePaused() Then
            updateEffect = updateGraviton(psEffect)
            Exit Function
        End If
        updateEffect = True
        Exit Function
    Case EFFECT_BLOOD
        If Not gamePaused() Then
            updateEffect = updateBlood(psEffect)
            Exit Function
        End If
        updateEffect = True
        Exit Function
    Case EFFECT_DESTRUCTION
        If Not gamePaused() Then
            updateEffect = updateDestruction(psEffect)
            Exit Function
        End If
        updateEffect = True
        Exit Function
    Case EFFECT_FIRE
        If Not gamePaused() Then
            updateEffect = updateFire(psEffect)
            Exit Function
        End If
        updateEffect = True
        Exit Function
    Case EFFECT_SAT_LASER
        If Not gamePaused() Then
            updateEffect = updateSatLaser(psEffect)
            Exit Function
        End If
        updateEffect = True
        Exit Function
    Case EFFECT_FIREWORK
        If Not gamePaused() Then
            updateEffect = updateFirework(psEffect)
            Exit Function
        End If
        updateEffect = True
        Exit Function
    Case EFFECT_DROID_ANIMEVENT_DYING
        updateEffect = updateDroidDeathAnimationEffect(psEffect)
        Exit Function
    Case EFFECT_FREED
    End Select
    debug_ 569, LOG_ERROR, "updateEffect", "Weirdy class of effect passed to updateEffect"
    exit_ 3                 ' abort()
End Function

' ----------------------------------------------------------------------------------------
' ALL THE UPDATE FUNCTIONS
' ----------------------------------------------------------------------------------------
'* Update the waypoint effects.
Private Function updateWaypoint(psEffect As EFFECT) As Boolean
    If Not (keyDown(KEY_LCTRL) Or keyDown(KEY_RCTRL) Or keyDown(KEY_LSHIFT) Or keyDown(KEY_RSHIFT)) Then
        updateWaypoint = False
        Exit Function
    End If
    updateWaypoint = True
End Function

Private Function updateFirework(psEffect As EFFECT) As Boolean
    Dim height As Long, xDif As Long, yDif As Long, radius As Long, val_ As Long
    Dim dv As Vector3i, dif As Long, drop As Long

    ' Move it
    psEffect.position.x = psEffect.position.x + graphicsTimeAdjustedIncrement(psEffect.velocity.x)
    psEffect.position.y = psEffect.position.y + graphicsTimeAdjustedIncrement(psEffect.velocity.y)
    psEffect.position.z = psEffect.position.z + graphicsTimeAdjustedIncrement(psEffect.velocity.z)

    If psEffect.typ = FIREWORK_TYPE_LAUNCHER Then
        height = ToInt(psEffect.position.y)
        If ULT(psEffect.size, height) Then
            dv.x = ToInt(psEffect.position.x)
            dv.z = ToInt(psEffect.position.z)
            dv.y = ToInt(psEffect.position.y + (psEffect.radius \ 2))
            addEffect dv, EFFECT_EXPLOSION, EXPLOSION_TYPE_MEDIUM, False, 0, 0, graphicsTime
            audio_PlayStaticTrack ToInt(psEffect.position.x), ToInt(psEffect.position.z), ID_SOUND_EXPLOSION

            For dif = 0 To psEffect.radius * 2 - 1 Step 20
                If dif < psEffect.radius Then
                    drop = psEffect.radius - dif
                Else
                    drop = dif - psEffect.radius
                End If
                radius = ToInt(Sqr(CSng(psEffect.radius * psEffect.radius - drop * drop)))
                'val = getStaticTimeValueRange(720,360);	// grab an angle - 4 seconds cyclic
                For val_ = 0 To 180 Step 20
                    ' (xDif, yDif are UDWORD: a negative sine or cosine becomes a huge float in the sums)
                    xDif = iSinR(DEG(val_), radius)
                    yDif = iCosR(DEG(val_), radius)
                    dv.x = ToInt(psEffect.position.x + CSng(U32(xDif)))
                    dv.z = ToInt(psEffect.position.z + CSng(U32(yDif)))
                    dv.y = ToInt(psEffect.position.y + dif)
                    effectGiveAuxVar 100
                    addEffect dv, EFFECT_FIREWORK, FIREWORK_TYPE_STARBURST, False, 0, 0, graphicsTime
                    dv.x = ToInt(psEffect.position.x - CSng(U32(xDif)))
                    dv.z = ToInt(psEffect.position.z - CSng(U32(yDif)))
                    dv.y = ToInt(psEffect.position.y + dif)
                    effectGiveAuxVar 100
                    addEffect dv, EFFECT_FIREWORK, FIREWORK_TYPE_STARBURST, False, 0, 0, graphicsTime
                    dv.x = ToInt(psEffect.position.x + CSng(U32(xDif)))
                    dv.z = ToInt(psEffect.position.z - CSng(U32(yDif)))
                    dv.y = ToInt(psEffect.position.y + dif)
                    effectGiveAuxVar 100
                    addEffect dv, EFFECT_FIREWORK, FIREWORK_TYPE_STARBURST, False, 0, 0, graphicsTime
                    dv.x = ToInt(psEffect.position.x - CSng(U32(xDif)))
                    dv.z = ToInt(psEffect.position.z + CSng(U32(yDif)))
                    dv.y = ToInt(psEffect.position.y + dif)
                    effectGiveAuxVar 100
                    addEffect dv, EFFECT_FIREWORK, FIREWORK_TYPE_STARBURST, False, 0, 0, graphicsTime
                Next
            Next
            updateFirework = False
            Exit Function
        Else
            ' Add an effect at the firework's position
            dv.x = ToInt(psEffect.position.x)
            dv.y = ToInt(psEffect.position.y)
            dv.z = ToInt(psEffect.position.z)

            ' Add a trail graphic
            addEffect dv, EFFECT_SMOKE, SMOKE_TYPE_TRAIL, False, 0, 0, graphicsTime
        End If
    Else    ' must be a startburst
        ' Time to update the frame number on the smoke sprite
        If ULT(psEffect.frameDelay, graphicsTime - psEffect.lastFrame) Then
            ' Store away last frame change time
            psEffect.lastFrame = graphicsTime

            ' Are we on the last frame?
            psEffect.frameNumber = (psEffect.frameNumber + 1) And &HFF&
            If Not ULT(psEffect.frameNumber, effectGetNumFrames(psEffect)) Then
                ' Does the anim wrap around?
                If (psEffect.control And EFFECT_CYCLIC) <> 0 Then
                    psEffect.frameNumber = 0
                Else
                    updateFirework = False ' Kill it off
                    Exit Function
                End If
            End If
        End If

        ' If it doesn't get killed by frame number, then by age
        If (psEffect.control And EFFECT_CYCLIC) <> 0 Then
            ' Has it overstayed it's welcome?
            If ULT(psEffect.lifeSpan, graphicsTime - psEffect.birthTime) Then
                updateFirework = False ' Kill it
                Exit Function
            End If
        End If
    End If
    updateFirework = True
End Function

Private Function updateSatLaser(psEffect As EFFECT) As Boolean
    Dim dv As Vector3i, val_ As Long, radius As Long, xDif As Long, yDif As Long, i As Long
    Dim startHeight As Long, endHeight As Long, xPos As Long, yPos As Long, lightPos As Vector3f

    ' Do these here because they are used by the lighting code below this if
    xPos = ToInt(psEffect.position.x)
    startHeight = ToInt(psEffect.position.y)
    endHeight = startHeight + 1064
    yPos = ToInt(psEffect.position.z)

    If psEffect.baseScale <> 0 Then
        psEffect.baseScale = 0

        ' Add some big explosions....!

        For i = 0 To 15
            dv.x = xPos + (200 - rand() Mod 400)
            dv.z = yPos + (200 - rand() Mod 400)
            dv.y = startHeight + rand() Mod 100
            addEffect dv, EFFECT_EXPLOSION, EXPLOSION_TYPE_MEDIUM, False, 0, 0, graphicsTime
        Next
        ' Add a sound effect
        audio_PlayStaticTrack ToInt(psEffect.position.x), ToInt(psEffect.position.z), ID_SOUND_EXPLOSION

        ' Add a shockwave
        dv.x = xPos
        dv.z = yPos
        dv.y = startHeight + SHOCK_WAVE_HEIGHT
        addEffect dv, EFFECT_EXPLOSION, EXPLOSION_TYPE_SHOCKWAVE, False, 0, 0, graphicsTime

        ' Now, add the column of light
        For i = startHeight To endHeight - 1 Step 56
            radius = 80
            ' Add 36 around in a circle..!
            For val_ = 0 To 180 Step 30
                xDif = iSinR(DEG(val_), radius)
                yDif = iCosR(DEG(val_), radius)
                dv.x = xPos + xDif + i \ 64
                dv.z = yPos + yDif
                dv.y = startHeight + i
                effectGiveAuxVar 100
                addEffect dv, EFFECT_EXPLOSION, EXPLOSION_TYPE_MEDIUM, False, 0, 0, graphicsTime
                dv.x = xPos - xDif + i \ 64
                dv.z = yPos - yDif
                dv.y = startHeight + i
                effectGiveAuxVar 100
                addEffect dv, EFFECT_EXPLOSION, EXPLOSION_TYPE_MEDIUM, False, 0, 0, graphicsTime
                dv.x = xPos + xDif + i \ 64
                dv.z = yPos - yDif
                dv.y = startHeight + i
                effectGiveAuxVar 100
                addEffect dv, EFFECT_EXPLOSION, EXPLOSION_TYPE_MEDIUM, False, 0, 0, graphicsTime
                dv.x = xPos - xDif + i \ 64
                dv.z = yPos + yDif
                dv.y = startHeight + i
                effectGiveAuxVar 100
                addEffect dv, EFFECT_EXPLOSION, EXPLOSION_TYPE_MEDIUM, False, 0, 0, graphicsTime
            Next
        Next
    End If

    If ULT(graphicsTime - psEffect.birthTime, 1000) Then
        lightPos = glm_vec3(CSng(xPos), CSng(startHeight), CSng(yPos))
        pushLight lightPos, 800, pal_Colour(0, 0, 255)
        updateSatLaser = True
    Else
        updateSatLaser = False
    End If
End Function

'* The update function for the explosions
Private Function updateExplosion(psEffect As EFFECT) As Boolean
    Dim percent As Long, range As Long, scaling As Single, lightRange As Long

    If (psEffect.control And EFFECT_LIT) <> 0 Then
        If psEffect.lifeSpan <> 0 Then
            percent = PERCENT_u(graphicsTime - psEffect.birthTime, psEffect.lifeSpan)
            If ULT(100, percent) Then
                percent = 100
            ElseIf percent > 50 Then
                percent = 100 - percent
            End If
        Else
            percent = 100
        End If

        If psEffect.typ <> EXPLOSION_TYPE_LAND_LIGHT Then
            range = percent
            pushLight psEffect.position, (3 * range) \ 2, pal_Colour(255, 0, 0)
        ElseIf Not rejectLandLight(psEffect.specific) Then
            range = percent
            pushLight psEffect.position, (3 * range) \ 2, pal_Colour(255, 0, 0)
        End If
    End If

    If psEffect.typ = EXPLOSION_TYPE_SHOCKWAVE Then
        scaling = CSng(psEffect.size) / CSng(MAX_SHOCKWAVE_SIZE)

        psEffect.size = (psEffect.size + (CLng(Fix(graphicsTimeAdjustedIncrement(CSng(SHOCKWAVE_SPEED)))) And &HFFFF&)) And &HFFFF&
        psEffect.frameNumber = CLng(Fix(scaling * effectGetNumFrames(psEffect))) And &HFF&

        lightRange = psEffect.size + 200
        pushLight psEffect.position, lightRange, pal_Colour(255, 255, 0)

        If psEffect.size > MAX_SHOCKWAVE_SIZE Or ULT(600, lightRange) Then
            updateExplosion = False ' Kill it off
            Exit Function
        End If
    ' Time to update the frame number on the explosion
    Else
        Do While ULT(psEffect.frameDelay, graphicsTime - psEffect.lastFrame)
            psEffect.lastFrame = psEffect.lastFrame + psEffect.frameDelay
            ' Are we on the last frame?

            psEffect.frameNumber = (psEffect.frameNumber + 1) And &HFF&
            If Not ULT(psEffect.frameNumber, effectGetNumFrames(psEffect)) Then
                If psEffect.typ <> EXPLOSION_TYPE_LAND_LIGHT Then
                    updateExplosion = False ' Kill it off
                    Exit Function
                Else
                    psEffect.frameNumber = 0
                End If
            End If
        Loop
    End If

    If Not gamePaused() Then
        ' Tesla explosions are the only ones that rise, or indeed move
        If psEffect.typ = EXPLOSION_TYPE_TESLA Then
            psEffect.position.y = psEffect.position.y + graphicsTimeAdjustedIncrement(psEffect.velocity.y)
        End If
    End If
    updateExplosion = True
End Function

'* The update function for blood
Private Function updateBlood(psEffect As EFFECT) As Boolean
    ' Time to update the frame number on the blood
    If ULT(psEffect.frameDelay, graphicsTime - psEffect.lastFrame) Then
        psEffect.lastFrame = graphicsTime
        ' Are we on the last frame?
        psEffect.frameNumber = (psEffect.frameNumber + 1) And &HFF&
        If Not ULT(psEffect.frameNumber, effectGetNumFrames(psEffect)) Then
            updateBlood = False ' Kill it off
            Exit Function
        End If
    End If
    ' Move it about in the world
    psEffect.position.x = psEffect.position.x + graphicsTimeAdjustedIncrement(psEffect.velocity.x)
    psEffect.position.y = psEffect.position.y + graphicsTimeAdjustedIncrement(psEffect.velocity.y)
    psEffect.position.z = psEffect.position.z + graphicsTimeAdjustedIncrement(psEffect.velocity.z)
    updateBlood = True
End Function

'* Processes all the drifting smoke. Handles the smoke puffing out the factory as well.
Private Function updatePolySmoke(psEffect As EFFECT) As Boolean
    ' Time to update the frame number on the smoke sprite
    Do While ULT(psEffect.frameDelay, graphicsTime - psEffect.lastFrame)
        ' Store away last frame change time
        psEffect.lastFrame = psEffect.lastFrame + psEffect.frameDelay

        ' Are we on the last frame?
        psEffect.frameNumber = (psEffect.frameNumber + 1) And &HFF&
        If Not ULT(psEffect.frameNumber, effectGetNumFrames(psEffect)) Then
            ' Does the anim wrap around?
            If (psEffect.control And EFFECT_CYCLIC) <> 0 Then
                ' Does it change drift direction?
                If psEffect.typ = SMOKE_TYPE_DRIFTING Then
                    ' Make it change direction
                    psEffect.velocity.x = CSng(rand() Mod 20)
                    psEffect.velocity.z = CSng(10 - rand() Mod 20)
                    psEffect.velocity.y = CSng(10 + rand() Mod 20)
                End If
                ' Reset the frame
                psEffect.frameNumber = 0
            Else
                ' Kill it off
                updatePolySmoke = False
                Exit Function
            End If
        End If
    Loop

    ' Update position
    psEffect.position.x = psEffect.position.x + graphicsTimeAdjustedIncrement(psEffect.velocity.x)
    psEffect.position.y = psEffect.position.y + graphicsTimeAdjustedIncrement(psEffect.velocity.y)
    psEffect.position.z = psEffect.position.z + graphicsTimeAdjustedIncrement(psEffect.velocity.z)

    ' If it doesn't get killed by frame number, then by age
    If (psEffect.control And EFFECT_CYCLIC) <> 0 Then
        ' Has it overstayed it's welcome?
        If ULT(psEffect.lifeSpan, graphicsTime - psEffect.birthTime) Then
            updatePolySmoke = False ' Kill it
            Exit Function
        End If
    End If
    updatePolySmoke = True
End Function

'
'   Gravitons just fly up for a bit and then drop down and are
'   killed off when they hit the ground
'
Private Function updateGraviton(psEffect As EFFECT) As Boolean
    Dim accel As Single, dv As Vector3i, psTile As Long, groundHeight As Long

    If psEffect.typ <> GRAVITON_TYPE_GIBLET Then
        pushLight psEffect.position, 128, pal_Colour(255, 255, 0)
    End If

    If gamePaused() Then
        ' Only update the lights if it's paused
        updateGraviton = True
        Exit Function
    End If

    ' Move it about in the world
    psEffect.position.x = psEffect.position.x + graphicsTimeAdjustedIncrement(psEffect.velocity.x)
    psEffect.position.y = psEffect.position.y + graphicsTimeAdjustedIncrement(psEffect.velocity.y)
    psEffect.position.z = psEffect.position.z + graphicsTimeAdjustedIncrement(psEffect.velocity.z)

    ' If it's bounced/drifted off the map then kill it
    If map_coord(ToInt(psEffect.position.x)) >= mapWidth Or map_coord(ToInt(psEffect.position.z)) >= mapHeight Then
        updateGraviton = False
        Exit Function
    End If

    groundHeight = map_Height(ToInt(psEffect.position.x), ToInt(psEffect.position.z))

    ' If it's going up and it's still under the landscape, then remove it...
    If psEffect.position.y < groundHeight And psEffect.velocity.y > 0 Then
        updateGraviton = False
        Exit Function
    End If

    ' Does it emit a trail? And is it high enough?
    If (psEffect.typ = GRAVITON_TYPE_EMITTING_DR) _
       Or ((psEffect.typ = GRAVITON_TYPE_EMITTING_ST) _
           And (psEffect.position.y > (groundHeight + 10))) Then
        ' Time to add another trail 'thing'?
        If ULT(psEffect.lastFrame + psEffect.frameDelay, graphicsTime) Then
            ' Store away last update
            psEffect.lastFrame = graphicsTime

            ' Add an effect at the gravitons's position
            dv.x = ToInt(psEffect.position.x)
            dv.y = ToInt(psEffect.position.y)
            dv.z = ToInt(psEffect.position.z)

            ' Add a trail graphic
            addEffect dv, EFFECT_SMOKE, SMOKE_TYPE_TRAIL, False, 0, 0, graphicsTime
        End If
    ElseIf psEffect.typ = GRAVITON_TYPE_GIBLET And (psEffect.position.y > (groundHeight + 5)) Then
        ' Time to add another trail 'thing'?
        If ULT(psEffect.lastFrame + psEffect.frameDelay, graphicsTime) Then
            ' Store away last update
            psEffect.lastFrame = graphicsTime

            ' Add an effect at the gravitons's position
            dv.x = ToInt(psEffect.position.x)
            dv.y = ToInt(psEffect.position.y)
            dv.z = ToInt(psEffect.position.z)
            addEffect dv, EFFECT_BLOOD, BLOOD_TYPE_NORMAL, False, 0, 0, graphicsTime
        End If
    End If

    ' Spin it round a bit
    psEffect.rotation.x = psEffect.rotation.x + ToInt(graphicsTimeAdjustedIncrement(CSng(psEffect.spin.x)))
    psEffect.rotation.y = psEffect.rotation.y + ToInt(graphicsTimeAdjustedIncrement(CSng(psEffect.spin.y)))
    psEffect.rotation.z = psEffect.rotation.z + ToInt(graphicsTimeAdjustedIncrement(CSng(psEffect.spin.z)))

    ' Update velocity (and retarding of descent) according to present frame rate
    accel = graphicsTimeAdjustedIncrement(GRAVITON_GRAVITY)
    psEffect.velocity.y = psEffect.velocity.y + accel

    ' If it's bounced/drifted off the map then kill it
    If ToInt(psEffect.position.x) <= TILE_UNITS Or ToInt(psEffect.position.z) <= TILE_UNITS Then
        updateGraviton = False
        Exit Function
    End If

    ' Are we below it? - Hit the ground?
    If ToInt(psEffect.position.y) < groundHeight Then
        psTile = mapTile(map_coord(ToInt(psEffect.position.x)), map_coord(ToInt(psEffect.position.z)))
        If terrainType(psTile) = TER_WATER Then
            updateGraviton = False
            Exit Function
        ElseIf ToInt(psEffect.velocity.y) < 0 Then ' Are we falling - rather than rising?
            ' Has it sufficient energy to keep bouncing?
            If Abs(psEffect.velocity.y) > 16 And psEffect.specific <= 2 Then
                psEffect.specific = (psEffect.specific + 1) And &HFF&

                ' Half it's velocity
                psEffect.velocity.y = psEffect.velocity.y / -2! ' only y gets flipped

                ' Also decrease other velocities
                psEffect.velocity.x = psEffect.velocity.x * 0.5!
                psEffect.velocity.z = psEffect.velocity.z * 0.5!

                ' Set it at ground level - may have gone through
                psEffect.position.y = CSng(groundHeight)
            Else
                ' Giblets don't blow up when they hit the ground!
                If psEffect.typ <> GRAVITON_TYPE_GIBLET Then
                    ' Remove the graviton and add an explosion
                    dv.x = ToInt(psEffect.position.x)
                    dv.y = ToInt(psEffect.position.y + 10)
                    dv.z = ToInt(psEffect.position.z)
                    addEffect dv, EFFECT_EXPLOSION, EXPLOSION_TYPE_VERY_SMALL, False, 0, 0, graphicsTime
                End If
                updateGraviton = False
                Exit Function
            End If
        End If
    End If
    updateGraviton = True
End Function


'* This isn't really an on-screen effect itself - it just spawns other ones....
Private Function updateDestruction(psEffect As EFFECT) As Boolean
    Dim pos As Vector3i, effectType As Long
    Dim widthScatter As Long, breadthScatter As Long, heightScatter As Long
    Dim iX As Long, iY As Long, percent As Long, range As Long, div As Single, height As Long
    Dim lightRange As Long, lightColour As Long, mx As Long, my As Long, mz As Long
    widthScatter = 1
    breadthScatter = 1
    heightScatter = 1

    percent = PERCENT_u(graphicsTime - psEffect.birthTime, psEffect.lifeSpan)
    If percent > 100 Then
        percent = 100
    End If
    range = 50 - Abs(50 - percent)

    If psEffect.typ = DESTRUCTION_TYPE_STRUCTURE Then
        lightRange = range * 10
    Else
        lightRange = range * 4
    End If
    If psEffect.typ = DESTRUCTION_TYPE_POWER_STATION Then
        lightRange = lightRange * 3
        lightColour = pal_Colour(255, 255, 255)
    Else
        lightColour = pal_Colour(255, 0, 0)
    End If
    pushLight psEffect.position, lightRange, lightColour

    If ULT(psEffect.birthTime + psEffect.lifeSpan, graphicsTime) Then
        ' Kill it - it's too old
        updateDestruction = False
        Exit Function
    End If

    If psEffect.typ = DESTRUCTION_TYPE_SKYSCRAPER Then

        If ULT((psEffect.lifeSpan * 9) \ 10, graphicsTime - psEffect.birthTime) Then
            pos.x = ToInt(psEffect.position.x)
            pos.z = ToInt(psEffect.position.z)
            pos.y = ToInt(psEffect.position.y)
            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_LARGE, False, 0, 0, graphicsTime
            updateDestruction = False
            Exit Function
        End If

        div = 1! - CSng(U32(graphicsTime - psEffect.birthTime)) / psEffect.lifeSpan
        If div < 0! Then
            div = 0!
        End If
        iIMDShape_max psEffect.imd, mx, my, mz
        height = ToInt(div * my)
    Else
        height = 16
    End If

    ' Time to add another effect?
    If ULT(psEffect.frameDelay, graphicsTime - psEffect.lastFrame) Then
        psEffect.lastFrame = graphicsTime
        Select Case psEffect.typ
        Case DESTRUCTION_TYPE_SKYSCRAPER
            widthScatter = TILE_UNITS
            breadthScatter = TILE_UNITS
            heightScatter = TILE_UNITS

        Case DESTRUCTION_TYPE_POWER_STATION, DESTRUCTION_TYPE_STRUCTURE
            widthScatter = TILE_UNITS \ 2
            breadthScatter = TILE_UNITS \ 2
            heightScatter = TILE_UNITS \ 4

        Case DESTRUCTION_TYPE_DROID, DESTRUCTION_TYPE_WALL_SECTION, DESTRUCTION_TYPE_FEATURE
            widthScatter = TILE_UNITS \ 6
            breadthScatter = TILE_UNITS \ 6
            heightScatter = TILE_UNITS \ 6
        Case Else
            ASSERT_FAILURE "false", "src/effects.cpp:1173", "updateDestruction", "Weirdy destruction type effect"
        End Select


        ' Find a position to dump it at
        pos.x = ToInt(psEffect.position.x + widthScatter - (rand() Mod (2 * widthScatter)))
        pos.z = ToInt(psEffect.position.z + breadthScatter - (rand() Mod (2 * breadthScatter)))
        pos.y = ToInt(psEffect.position.y + height + (rand() Mod heightScatter))

        If psEffect.typ = DESTRUCTION_TYPE_SKYSCRAPER Then
            pos.y = ToInt(psEffect.position.y + height)
        End If


        ' Choose an effect
        effectType = rand() Mod 15
        Select Case effectType
        Case 0
            addEffect pos, EFFECT_SMOKE, SMOKE_TYPE_DRIFTING, False, 0, 0, graphicsTime
        Case 1, 2, 3, 4, 5
            If psEffect.typ = DESTRUCTION_TYPE_SKYSCRAPER Then
                addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_LARGE, False, 0, 0, graphicsTime
            ' Only structures get the big explosions
            ElseIf psEffect.typ = DESTRUCTION_TYPE_STRUCTURE Then
                addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_MEDIUM, False, 0, 0, graphicsTime
            Else
                addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_SMALL, False, 0, 0, graphicsTime
            End If
        Case 6, 7, 8, 9, 10
            If psEffect.typ = DESTRUCTION_TYPE_STRUCTURE Then
                addEffect pos, EFFECT_GRAVITON, GRAVITON_TYPE_EMITTING_ST, True, iIMDBaseShape_displayModel(getRandomDebrisImd()), 0, graphicsTime
            Else
                addEffect pos, EFFECT_GRAVITON, GRAVITON_TYPE_EMITTING_DR, True, iIMDBaseShape_displayModel(getRandomDebrisImd()), 0, graphicsTime
            End If
        Case 11
            addEffect pos, EFFECT_SMOKE, SMOKE_TYPE_DRIFTING, False, 0, 0, graphicsTime
        Case 12, 13
            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_SMALL, False, 0, 0, graphicsTime
        Case 14
            ' Add sound effect, but only if we're less than 3/4 of the way thru' destruction
            If ULT(graphicsTime, FromU32(Int(U32(MUL32(psEffect.birthTime + psEffect.lifeSpan, 3)) / 4#))) Then
                iX = ToInt(psEffect.position.x)
                iY = ToInt(psEffect.position.z)
                audio_PlayStaticTrack iX, iY, ID_SOUND_EXPLOSION
            End If

        End Select
    End If
    updateDestruction = True
End Function

'* Moves the construction graphic about - dust cloud or whatever....
Private Function updateConstruction(psEffect As EFFECT) As Boolean
    ' Time to update the frame number on the construction sprite
    If ULT(psEffect.frameDelay, graphicsTime - psEffect.lastFrame) Then
        psEffect.lastFrame = graphicsTime
        ' Are we on the last frame?
        psEffect.frameNumber = (psEffect.frameNumber + 1) And &HFF&
        If Not ULT(psEffect.frameNumber, effectGetNumFrames(psEffect)) Then
            ' Is it a cyclic sprite?
            If (psEffect.control And EFFECT_CYCLIC) <> 0 Then
                psEffect.frameNumber = 0
            Else
                updateConstruction = False
                Exit Function
            End If
        End If
    End If

    ' Move it about in the world

    psEffect.position.x = psEffect.position.x + graphicsTimeAdjustedIncrement(psEffect.velocity.x)
    psEffect.position.y = psEffect.position.y + graphicsTimeAdjustedIncrement(psEffect.velocity.y)
    psEffect.position.z = psEffect.position.z + graphicsTimeAdjustedIncrement(psEffect.velocity.z)


    ' If it doesn't get killed by frame number, then by height
    If (psEffect.control And EFFECT_CYCLIC) <> 0 Then
        ' Has it hit the ground
        If ToInt(psEffect.position.y) <= map_Height(ToInt(psEffect.position.x), ToInt(psEffect.position.z)) Then
            updateConstruction = False
            Exit Function
        End If

        If ULT(psEffect.lifeSpan, graphicsTime - psEffect.birthTime) Then
            updateConstruction = False
            Exit Function
        End If
    End If
    updateConstruction = True
End Function

'* Update fire sequences
Private Function updateFire(psEffect As EFFECT) As Boolean
    Const fireLightHeightAboveGround As Long = 28
    Dim pos As Vector3i, percent As Long, lifeSpan1 As Long, lightPos As Vector3f, r1 As Long, r2 As Long

    If psEffect.lifeSpan > 1 Then lifeSpan1 = psEffect.lifeSpan Else lifeSpan1 = 1
    percent = PERCENT_u(graphicsTime - psEffect.birthTime, lifeSpan1)
    If ULT(100, percent) Then
        percent = 100
    End If

    lightPos = psEffect.position
    ' light.position = psEffect->position; light.position.y += fireLightHeightAboveGround (on the Vector3i)
    LightingData_lights_push_back ToInt(lightPos.x), ToInt(lightPos.y) + fireLightHeightAboveGround, ToInt(lightPos.z), _
                                  (percent * psEffect.radius * 6) \ 100, pal_Colour(255, 0, 0)

    ' Time to update the frame number on the construction sprite
    If ULT(psEffect.frameDelay, graphicsTime - psEffect.lastFrame) Then
        psEffect.lastFrame = graphicsTime
        r1 = rand() Mod psEffect.radius
        r2 = rand() Mod (2 * psEffect.radius)
        pos.x = ToInt(psEffect.position.x + (r1 - r2))
        r1 = rand() Mod psEffect.radius
        r2 = rand() Mod (2 * psEffect.radius)
        pos.z = ToInt(psEffect.position.z + (r1 - r2))

        ' Effect is off map, no need to update it anymore
        If Not worldOnMap(pos.x, pos.z) Then
            updateFire = False
            Exit Function
        End If

        pos.y = map_Height(pos.x, pos.z)

        If psEffect.typ = FIRE_TYPE_SMOKY_BLUE Then
            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_FLAMETHROWER, False, 0, 0, graphicsTime
        Else
            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_SMALL, False, 0, 0, graphicsTime
        End If

        If psEffect.typ = FIRE_TYPE_SMOKY Or psEffect.typ = FIRE_TYPE_SMOKY_BLUE Then
            r1 = (rand() Mod psEffect.radius) \ 2
            r2 = rand() Mod ((2 * psEffect.radius) \ 2)
            pos.x = ToInt(psEffect.position.x + (r1 - r2))
            r1 = (rand() Mod psEffect.radius) \ 2
            r2 = rand() Mod ((2 * psEffect.radius) \ 2)
            pos.z = ToInt(psEffect.position.z + (r1 - r2))
            pos.y = map_Height(pos.x, pos.z)
            addEffect pos, EFFECT_SMOKE, SMOKE_TYPE_DRIFTING_HIGH, False, 0, 0, graphicsTime
        Else
            r1 = rand() Mod psEffect.radius
            r2 = rand() Mod (2 * psEffect.radius)
            pos.x = ToInt(psEffect.position.x + (r1 - r2))
            r1 = rand() Mod psEffect.radius
            r2 = rand() Mod (2 * psEffect.radius)
            pos.z = ToInt(psEffect.position.z + (r1 - r2))

            ' Effect is off map, no need to update it anymore
            If Not worldOnMap(pos.x, pos.z) Then
                updateFire = False
                Exit Function
            End If

            pos.y = map_Height(pos.x, pos.z)
            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_SMALL, False, 0, 0, graphicsTime
        End If

        '
        'pos.x = psEffect->position.x;
        'pos.y = psEffect->position.y;
        'pos.z = psEffect->position.z;
        '
        'scatter.x = psEffect->radius; scatter.y = 0; scatter.z = psEffect->radius;
        'addMultiEffect(&pos,&scatter,EFFECT_EXPLOSION,EXPLOSION_TYPE_SMALL,false,NULL,2,0,0);
        '

    End If

    If ULT(psEffect.lifeSpan, graphicsTime - psEffect.birthTime) Then
        updateFire = False
        Exit Function
    End If
    updateFire = True
End Function

' ----------------------------------------------------------------------------------------
' ALL THE RENDER FUNCTIONS
' ----------------------------------------------------------------------------------------
'* Calls the appropriate render routine for each type of effect (psEffect: an EFFECT handle)
Public Sub renderEffect(ByVal psEffect As Long, viewMatrix As mat4)
    ' What type of effect are we dealing with?
    Select Case aEffect(psEffect).group
    Case EFFECT_WAYPOINT
        renderWaypointEffect aEffect(psEffect), viewMatrix
        Exit Sub

    Case EFFECT_EXPLOSION
        renderExplosionEffect aEffect(psEffect), viewMatrix
        Exit Sub

    Case EFFECT_CONSTRUCTION
        renderConstructionEffect aEffect(psEffect), viewMatrix
        Exit Sub

    Case EFFECT_SMOKE
        renderSmokeEffect aEffect(psEffect), viewMatrix
        Exit Sub

    Case EFFECT_GRAVITON
        renderGravitonEffect aEffect(psEffect), viewMatrix
        Exit Sub

    Case EFFECT_BLOOD
        renderBloodEffect aEffect(psEffect), viewMatrix
        Exit Sub

    Case EFFECT_DESTRUCTION
        '   There is no display func for a destruction effect -
        '   it merely spawn other effects over time
        renderDestructionEffect aEffect(psEffect), viewMatrix
        Exit Sub

    Case EFFECT_FIRE
        ' Likewise
        Exit Sub

    Case EFFECT_SAT_LASER
        ' Likewise
        Exit Sub

    Case EFFECT_FIREWORK
        renderFirework aEffect(psEffect), viewMatrix
        Exit Sub

    Case EFFECT_DROID_ANIMEVENT_DYING
        renderDroidDeathAnimationEffect aEffect(psEffect), viewMatrix
        Exit Sub

    Case EFFECT_FREED
    End Select

    debug_ 1441, LOG_ERROR, "renderEffect", "Weirdy class of effect passed to renderEffect"
    exit_ 3                 ' abort()
End Sub

'* drawing func for wapypoints
Private Sub renderWaypointEffect(psEffect As EFFECT, viewMatrix As mat4)
    Dim modelMatrix As mat4
    modelMatrix = positionEffect(psEffect)
    pie_Draw3DShape psEffect.imd, 0, 0, psPalette(WZCOL_WHITE), 0, 0, VarPtr(modelMatrix), VarPtr(viewMatrix)
End Sub

' * glm::rotate(UNDEG(-playerPos.r.y), y) * glm::rotate(UNDEG(-playerPos.r.x), x): the "face the viewer" rotations
Private Function faceViewer() As mat4
    faceViewer = glm_mat4_mul(glm_rotate(UNDEG(-playerPos.r.y), glm_vec3(0!, 1!, 0!)), glm_rotate(UNDEG(-playerPos.r.x), glm_vec3(1!, 0!, 0!)))
End Function

Private Sub renderFirework(psEffect As EFFECT, viewMatrix As mat4)
    Dim modelMatrix As mat4, s As Single
    ' these don't get rendered
    If psEffect.typ = FIREWORK_TYPE_LAUNCHER Then
        Exit Sub
    End If

    modelMatrix = positionEffect(psEffect)
    s = psEffect.size / 100!
    modelMatrix = glm_mat4_mul(modelMatrix, glm_mat4_mul(faceViewer(), glm_scale(glm_vec3(s, s, s))))

    pie_Draw3DShape psEffect.imd, psEffect.frameNumber, 0, psPalette(WZCOL_WHITE), pie_ADDITIVE Or pie_NODEPTHWRITE, EFFECT_EXPLOSION_ADDITIVE, VarPtr(modelMatrix), VarPtr(viewMatrix)
End Sub

'* drawing func for blood.
Private Sub renderBloodEffect(psEffect As EFFECT, viewMatrix As mat4)
    Dim modelMatrix As mat4, s As Single
    modelMatrix = positionEffect(psEffect)
    s = psEffect.size / 100!
    modelMatrix = glm_mat4_mul(modelMatrix, glm_mat4_mul(faceViewer(), glm_scale(glm_vec3(s, s, s))))

    pie_Draw3DShape getDisplayImdFromIndex(MI_BLOOD), psEffect.frameNumber, 0, psPalette(WZCOL_WHITE), pie_TRANSLUCENT Or pie_NODEPTHWRITE, EFFECT_BLOOD_TRANSPARENCY, VarPtr(modelMatrix), VarPtr(viewMatrix)
End Sub

Private Sub renderDestructionEffect(psEffect As EFFECT, viewMatrix As mat4)
    Dim div As Single, percent As Long, modelMatrix As mat4, sx As Long, sy As Long, sz As Long

    If psEffect.typ <> DESTRUCTION_TYPE_SKYSCRAPER Then
        Exit Sub
    End If

    modelMatrix = positionEffect(psEffect)

    div = CSng(U32(graphicsTime - psEffect.birthTime)) / psEffect.lifeSpan
    If div > 1! Then
        div = 1!    'temporary!
    End If
    div = 1! - div
    percent = Fix(div * pie_RAISE_SCALE)

    If Not gamePaused() Then
        ' (the three SKY_SHIMMY rand() calls, in the order of the rotations)
        sx = SKY_SHIMMY()
        sy = SKY_SHIMMY()
        sz = SKY_SHIMMY()
        modelMatrix = glm_mat4_mul(modelMatrix, glm_mat4_mul(glm_mat4_mul( _
            glm_rotate(UNDEG(sx), glm_vec3(1!, 0!, 0!)), _
            glm_rotate(UNDEG(sy), glm_vec3(0!, 1!, 0!))), _
            glm_rotate(UNDEG(sz), glm_vec3(0!, 0!, 1!))))
    End If
    pie_Draw3DShape psEffect.imd, 0, 0, psPalette(WZCOL_WHITE), pie_RAISE, percent, VarPtr(modelMatrix), VarPtr(viewMatrix)
End Sub

Private Function rejectLandLight(ByVal typ As LAND_LIGHT_SPEC) As Boolean
    Dim timeSlice As Long
    timeSlice = FromU32(U32(graphicsTime) - Int(U32(graphicsTime) / 2000#) * 2000#)

    If timeSlice < 400 Then
        rejectLandLight = (typ <> LL_MIDDLE) ' reject all expect middle
    ElseIf timeSlice < 800 Then
        rejectLandLight = (typ = LL_OUTER) ' reject only outer
    ElseIf timeSlice < 1200 Then
        rejectLandLight = False 'reject none
    ElseIf timeSlice < 1600 Then
        rejectLandLight = (typ = LL_OUTER) ' reject only outer
    Else
        rejectLandLight = (typ <> LL_MIDDLE) ' reject all expect middle
    End If
End Function

'* Renders the standard explosion effect
Private Sub renderExplosionEffect(psEffect As EFFECT, viewMatrix As mat4)
    Dim brightness As Long, modelMatrix As mat4, scale_ As Single, s As Single, premultiplied As Boolean
    brightness = psPalette(WZCOL_WHITE)

    If psEffect.imd = 0 Then
        Exit Sub
    End If

    If psEffect.typ = EXPLOSION_TYPE_LAND_LIGHT Then
        If rejectLandLight(psEffect.specific) Then
            Exit Sub
        End If
    End If

    modelMatrix = positionEffect(psEffect)

    ' Bit in comments - doesn't quite work yet?
    If (psEffect.control And EFFECT_FACING) <> 0 Then
        ' Always face the viewer!
        ' TODO This only faces towards the viewer, if the effect is in the middle of the screen... It draws the effect parallel with the screens near/far planes.
        modelMatrix = glm_mat4_mul(modelMatrix, faceViewer())
    End If

    ' Tesla explosions diminish in size
    If psEffect.typ = EXPLOSION_TYPE_TESLA Then
        scale_ = CSng(U32(graphicsTime - psEffect.birthTime)) / CSng(psEffect.lifeSpan)
        If scale_ < 0! Then
            scale_ = 0!
        ElseIf scale_ > 0.45! Then
            scale_ = 0.45!
        End If
        s = psEffect.size / 100! - scale_
        modelMatrix = glm_mat4_mul(modelMatrix, glm_scale(glm_vec3(s, s, s)))
    ElseIf psEffect.typ = EXPLOSION_TYPE_PLASMA Then
        scale_ = CSng(U32(graphicsTime - psEffect.birthTime)) / CSng(psEffect.lifeSpan) / 3!
        s = BASE_PLASMA_SIZE / 100! + scale_
        modelMatrix = glm_mat4_mul(modelMatrix, glm_scale(glm_vec3(s, s, s)))
    Else
        s = psEffect.size / 100!
        modelMatrix = glm_mat4_mul(modelMatrix, glm_scale(glm_vec3(s, s, s)))
    End If

    premultiplied = False
    If (iIMDShape_flags(psEffect.imd) And iV_IMD_PREMULTIPLIED) <> 0 Then
        premultiplied = True
    End If

    If premultiplied Then
        pie_Draw3DShape psEffect.imd, psEffect.frameNumber, 0, brightness, pie_PREMULTIPLIED Or pie_NODEPTHWRITE, 0, VarPtr(modelMatrix), VarPtr(viewMatrix)
    ElseIf psEffect.typ = EXPLOSION_TYPE_PLASMA Then
        pie_Draw3DShape psEffect.imd, psEffect.frameNumber, 0, brightness, pie_ADDITIVE Or pie_NODEPTHWRITE, EFFECT_PLASMA_ADDITIVE, VarPtr(modelMatrix), VarPtr(viewMatrix)
    ElseIf psEffect.typ = EXPLOSION_TYPE_KICKUP Then
        pie_Draw3DShape psEffect.imd, psEffect.frameNumber, 0, brightness, pie_TRANSLUCENT Or pie_NODEPTHWRITE, 128, VarPtr(modelMatrix), VarPtr(viewMatrix)
    Else
        pie_Draw3DShape psEffect.imd, psEffect.frameNumber, 0, brightness, pie_ADDITIVE Or pie_NODEPTHWRITE, EFFECT_EXPLOSION_ADDITIVE, VarPtr(modelMatrix), VarPtr(viewMatrix)
    End If
End Sub

Private Sub renderGravitonEffect(psEffect As EFFECT, viewMatrix As mat4)
    Dim modelMatrix As mat4, s As Single
    modelMatrix = positionEffect(psEffect)
    modelMatrix = glm_mat4_mul(modelMatrix, glm_mat4_mul(glm_mat4_mul( _
        glm_rotate(UNDEG(psEffect.rotation.x), glm_vec3(1!, 0!, 0!)), _
        glm_rotate(UNDEG(psEffect.rotation.y), glm_vec3(0!, 1!, 0!))), _
        glm_rotate(UNDEG(psEffect.rotation.z), glm_vec3(0!, 0!, 1!))))

    ' Buildings emitted by gravitons are chunkier
    If psEffect.typ = GRAVITON_TYPE_EMITTING_ST Then
        ' Twice as big - 150 percent
        s = psEffect.size / 100!
        modelMatrix = glm_mat4_mul(modelMatrix, glm_scale(glm_vec3(s, s, s)))
    End If

    pie_Draw3DShape psEffect.imd, psEffect.frameNumber, psEffect.player, psPalette(WZCOL_WHITE), pie_SHADOW, 0, VarPtr(modelMatrix), VarPtr(viewMatrix)
End Sub

'* Renders the standard construction effect
Private Sub renderConstructionEffect(psEffect As EFFECT, viewMatrix As mat4)
    Dim null_ As Vector3i, percent As Long, translucency As Long, size As Single, modelMatrix As mat4

    ' No rotation about arbitrary axis
    null_.x = 0
    null_.y = 0
    null_.z = 0

    modelMatrix = positionEffect(psEffect)

    ' Bit in comments doesn't quite work yet?
    If (psEffect.control And EFFECT_FACING) <> 0 Then
        modelMatrix = glm_mat4_mul(modelMatrix, faceViewer())
    End If

    ' Scale size according to age
    percent = PERCENT_u(graphicsTime - psEffect.birthTime, psEffect.lifeSpan)
    If percent < 0 Then
        percent = 0
    End If
    If percent > 100 Then
        percent = 100
    End If

    ' Make imds be transparent on 3dfx
    If percent < 50 Then
        translucency = percent * 2
    Else
        translucency = (100 - percent) * 2
    End If
    translucency = translucency + 10
    size = 2! * translucency / 100!
    If size > 0.9! Then size = 0.9!
    modelMatrix = glm_mat4_mul(modelMatrix, glm_scale(glm_vec3(size, size, size)))

    pie_Draw3DShape psEffect.imd, psEffect.frameNumber, 0, psPalette(WZCOL_WHITE), pie_TRANSLUCENT Or pie_NODEPTHWRITE, translucency, VarPtr(modelMatrix), VarPtr(viewMatrix)
End Sub

'* Renders the standard smoke effect - it is now scaled in real-time as well
Private Sub renderSmokeEffect(psEffect As EFFECT, viewMatrix As mat4)
    Dim transparency As Long, brightness As Long, modelMatrix As mat4
    Dim lifetime As Long, percent As Long, scale_ As Single
    transparency = 0
    brightness = psPalette(WZCOL_WHITE)

    modelMatrix = positionEffect(psEffect)

    ' Bit in comments doesn't quite work yet?
    If (psEffect.control And EFFECT_FACING) <> 0 Then
        ' Always face the viewer!
        modelMatrix = glm_mat4_mul(modelMatrix, faceViewer())
    End If

    If (psEffect.control And EFFECT_SCALED) <> 0 Then
        lifetime = graphicsTime - psEffect.birthTime

        ' Since psEffect->birthTime will be set to graphicsTime on
        ' creation, and graphicsTime only increments, birthTime should be
        ' smaller than or equal to graphicsTime. As a great man once said
        ' though (hehe, I just love exaggerating ;-) -- Giel):
        ' "Always assert your assumptions!". So here it goes:
        ' (assert(psEffect->birthTime <= graphicsTime): the C assert, off in a release build)

        If Not (psEffect.lifeSpan <> 0) Then ASSERT_FAILURE "psEffect->lifeSpan != 0", "src/effects.cpp:1705", "renderSmokeEffect", "Effect lifespan is zero (seconds); it really should be non-zero!"

        ' HACK: Work around a bug that bit me causing a "integer divide by zero" at this location -- Giel
        If psEffect.lifeSpan <> 0 Then
            percent = PERCENT_u(lifetime, psEffect.lifeSpan)
        Else
            percent = 100
        End If

        scale_ = (percent + psEffect.baseScale) / 100!
        modelMatrix = glm_mat4_mul(modelMatrix, glm_scale(glm_vec3(scale_, scale_, scale_)))
        transparency = FromU32(Int(U32(MUL32(EFFECT_SMOKE_TRANSPARENCY, 100 - percent)) / 100#))
    End If

    transparency = (transparency * 3) \ 2  'JPS smoke strength increased for d3d 12 may 99

    ' Make imds be transparent on 3dfx
    If psEffect.typ = SMOKE_TYPE_STEAM Then
        pie_Draw3DShape psEffect.imd, psEffect.frameNumber, 0, brightness, pie_TRANSLUCENT Or pie_NODEPTHWRITE, EFFECT_STEAM_TRANSPARENCY \ 2, VarPtr(modelMatrix), VarPtr(viewMatrix)
    Else
        If psEffect.typ = SMOKE_TYPE_TRAIL Then
            pie_Draw3DShape psEffect.imd, psEffect.frameNumber, 0, brightness, pie_TRANSLUCENT Or pie_NODEPTHWRITE, (2 * transparency) \ 3, VarPtr(modelMatrix), VarPtr(viewMatrix)
        Else
            pie_Draw3DShape psEffect.imd, psEffect.frameNumber, 0, brightness, pie_TRANSLUCENT Or pie_NODEPTHWRITE, transparency \ 2, VarPtr(modelMatrix), VarPtr(viewMatrix)
        End If
    End If
End Sub


' ----------------------------------------------------------------------------------------
' ALL THE SETUP FUNCTIONS
' ----------------------------------------------------------------------------------------
Private Sub effectSetupFirework(effect_ As EFFECT)
    If effect_.typ = FIREWORK_TYPE_LAUNCHER Then
        effect_.velocity.x = 200 - rand() Mod 400
        effect_.velocity.z = 200 - rand() Mod 400
        effect_.velocity.y = 400 + rand() Mod 200   'height
        effect_.lifeSpan = GAME_TICKS_PER_SEC * 3
        effect_.radius = 80 + rand() Mod 150
        effect_.size = 300 + rand() Mod 300 'height it goes off
        effect_.imd = getDisplayImdFromIndex(MI_FIREWORK) ' not actually drawn
    Else
        effect_.velocity.x = 20 - rand() Mod 40
        effect_.velocity.z = 20 - rand() Mod 40
        effect_.velocity.y = 0 - (20 + rand() Mod 40)   'height
        effect_.lifeSpan = GAME_TICKS_PER_SEC * 4

        ' setup the imds
        Select Case rand() Mod 3
        Case 0
            effect_.imd = getDisplayImdFromIndex(MI_FIREWORK)
            effect_.size = 45   'size of graphic
        Case 1
            effect_.imd = getDisplayImdFromIndex(MI_SNOW)
            effect_.control = effect_.control Or EFFECT_CYCLIC
            effect_.size = 60   'size of graphic

        Case Else
            effect_.imd = getDisplayImdFromIndex(MI_FLAME)
            effect_.size = 40   'size of graphic


        End Select
    End If
    effect_.frameDelay = (EXPLOSION_FRAME_DELAY() * 2) And &HFFFF&
End Sub

Private Sub effectSetupSmoke(effect_ As EFFECT)
    ' everything except steam drifts about
    If effect_.typ = SMOKE_TYPE_STEAM Then
        ' Only upwards
        effect_.velocity.x = 0!
        effect_.velocity.z = 0!
    ElseIf effect_.typ = SMOKE_TYPE_BILLOW Then

        effect_.velocity.x = CSng(10 - rand() Mod 20)
        effect_.velocity.z = CSng(10 - rand() Mod 20)
    Else
        effect_.velocity.x = CSng(rand() Mod 20)
        effect_.velocity.z = CSng(10 - rand() Mod 20)
    End If

    ' Steam isn't cyclic  - it doesn't grow with time either
    If effect_.typ <> SMOKE_TYPE_STEAM Then
        effect_.control = effect_.control Or EFFECT_CYCLIC
        effect_.control = effect_.control Or EFFECT_SCALED
    End If

    Select Case effect_.typ
    Case SMOKE_TYPE_DRIFTING
        effect_.imd = getDisplayImdFromIndex(MI_SMALL_SMOKE)
        effect_.lifeSpan = NORMAL_SMOKE_LIFESPAN() And &HFFFF&
        effect_.velocity.y = CSng(35 + rand() Mod 30)
        effect_.baseScale = 40
    Case SMOKE_TYPE_DRIFTING_HIGH
        effect_.imd = getDisplayImdFromIndex(MI_SMALL_SMOKE)
        effect_.lifeSpan = NORMAL_SMOKE_LIFESPAN() And &HFFFF&
        effect_.velocity.y = CSng(40 + rand() Mod 45)
        effect_.baseScale = 25
    Case SMOKE_TYPE_DRIFTING_SMALL
        effect_.imd = getDisplayImdFromIndex(MI_SMALL_SMOKE)
        effect_.lifeSpan = SMALL_SMOKE_LIFESPAN() And &HFFFF&
        effect_.velocity.y = CSng(25 + rand() Mod 35)
        effect_.baseScale = 17
    Case SMOKE_TYPE_BILLOW
        effect_.imd = getDisplayImdFromIndex(MI_SMALL_SMOKE)
        effect_.lifeSpan = SMALL_SMOKE_LIFESPAN() And &HFFFF&
        effect_.velocity.y = CSng(10 + rand() Mod 20)
        effect_.baseScale = 80
    Case SMOKE_TYPE_STEAM
        effect_.imd = getDisplayImdFromIndex(MI_SMALL_STEAM)
        effect_.velocity.y = CSng(rand() Mod 5)
    Case SMOKE_TYPE_TRAIL
        effect_.imd = getDisplayImdFromIndex(MI_TRAIL)
        effect_.lifeSpan = TRAIL_SMOKE_LIFESPAN
        effect_.velocity.y = CSng(5 + rand() Mod 10)
        effect_.baseScale = 25
    Case Else
        ASSERT_FAILURE "false", "src/effects.cpp:1854", "effectSetupSmoke", "Weird smoke type"
    End Select

    ' It always faces you
    effect_.control = effect_.control Or EFFECT_FACING

    effect_.frameDelay = SMOKE_FRAME_DELAY() And &HFFFF&
    ' Randomly flip gfx for variation
    If ONEINTWO() Then
        effect_.control = effect_.control Or EFFECT_X_FLIP
    End If
    If ONEINTWO() Then
        effect_.control = effect_.control Or EFFECT_Y_FLIP
    End If
End Sub

Private Sub effectSetupSatLaser(effect_ As EFFECT)
    ' Does nothing at all..... Runs only for one frame!
    effect_.baseScale = 1
End Sub

Private Sub effectSetupGraviton(effect_ As EFFECT)
    Select Case effect_.typ
    Case GRAVITON_TYPE_GIBLET
        effect_.velocity.x = effect_.velocity.x + GIBLET_INIT_VEL_X()
        effect_.velocity.z = effect_.velocity.z + GIBLET_INIT_VEL_Z()
        effect_.velocity.y = GIBLET_INIT_VEL_Y
    Case GRAVITON_TYPE_EMITTING_ST
        effect_.velocity.x = GRAVITON_INIT_VEL_X()
        effect_.velocity.z = GRAVITON_INIT_VEL_Z()
        effect_.velocity.y = (5 * GRAVITON_INIT_VEL_Y()) / 4
        effect_.size = (120 + rand() Mod 30) And &HFFFF&
    Case GRAVITON_TYPE_EMITTING_DR
        effect_.velocity.x = effect_.velocity.x + GRAVITON_INIT_VEL_X() / 2
        effect_.velocity.z = effect_.velocity.z + GRAVITON_INIT_VEL_Z() / 2
        effect_.velocity.y = GRAVITON_INIT_VEL_Y()
    Case Else
        ASSERT_FAILURE "false", "src/effects.cpp:1901", "effectSetupGraviton", "Weirdy type of graviton"

    End Select

    effect_.rotation.x = rand() Mod DEG(360)
    effect_.rotation.z = rand() Mod DEG(360)
    effect_.rotation.y = rand() Mod DEG(360)

    effect_.spin.x = rand() Mod DEG(100) + DEG(20)
    effect_.spin.z = rand() Mod DEG(100) + DEG(20)
    effect_.spin.y = rand() Mod DEG(100) + DEG(20)

    ' Gravitons are essential
    effect_.control = effect_.control Or EFFECT_ESSENTIAL

    If effect_.typ = GRAVITON_TYPE_GIBLET Then
        effect_.frameDelay = GRAVITON_BLOOD_DELAY() And &HFFFF&
    Else
        effect_.frameDelay = GRAVITON_FRAME_DELAY() And &HFFFF&
    End If
End Sub

Private Sub effectSetupExplosion(effect_ As EFFECT)
    ' Get an imd if it's not established
    If effect_.imd = 0 Then
        Select Case effect_.typ
        Case EXPLOSION_TYPE_SMALL
            effect_.imd = getDisplayImdFromIndex(MI_EXPLOSION_SMALL)
            effect_.size = ((6 * EXPLOSION_SIZE()) \ 5) And &HFF&
        Case EXPLOSION_TYPE_VERY_SMALL
            effect_.imd = getDisplayImdFromIndex(MI_EXPLOSION_SMALL)
            effect_.size = (BASE_FLAME_SIZE + auxVar) And &HFF&
        Case EXPLOSION_TYPE_MEDIUM
            effect_.imd = getDisplayImdFromIndex(MI_EXPLOSION_MEDIUM)
            effect_.size = EXPLOSION_SIZE() And &HFF&
        Case EXPLOSION_TYPE_LARGE
            effect_.imd = getDisplayImdFromIndex(MI_EXPLOSION_MEDIUM)
            effect_.size = ((EXPLOSION_SIZE() And &HFF&) * 2) And &HFFFF&
        Case EXPLOSION_TYPE_FLAMETHROWER
            effect_.imd = getDisplayImdFromIndex(MI_FLAME)
            effect_.size = (BASE_FLAME_SIZE + auxVar) And &HFF&
        Case EXPLOSION_TYPE_LASER
            effect_.imd = getDisplayImdFromIndex(MI_FLAME)  ' change this
            effect_.size = (BASE_LASER_SIZE + auxVar) And &HFF&
        Case EXPLOSION_TYPE_DISCOVERY
            effect_.imd = getDisplayImdFromIndex(MI_TESLA)  ' change this
            effect_.size = DISCOVERY_SIZE
        Case EXPLOSION_TYPE_FLARE
            effect_.imd = getDisplayImdFromIndex(MI_MFLARE)
            effect_.size = FLARE_SIZE
        Case EXPLOSION_TYPE_TESLA
            effect_.imd = getDisplayImdFromIndex(MI_TESLA)
            effect_.size = TESLA_SIZE
            effect_.velocity.y = CSng(TESLA_SPEED)

        Case EXPLOSION_TYPE_KICKUP
            effect_.imd = getDisplayImdFromIndex(MI_KICK)
            effect_.size = 100
        Case EXPLOSION_TYPE_PLASMA
            effect_.imd = getDisplayImdFromIndex(MI_PLASMA)
            effect_.size = BASE_PLASMA_SIZE
            effect_.velocity.y = 0!
        Case EXPLOSION_TYPE_LAND_LIGHT
            effect_.imd = getDisplayImdFromIndex(MI_LANDING)
            effect_.size = 120
            effect_.specific = ellSpec And &HFF&
            effect_.velocity.y = 0!
            effect_.control = effect_.control Or EFFECT_ESSENTIAL ' Landing lights are permanent and cyclic
        Case EXPLOSION_TYPE_SHOCKWAVE
            effect_.imd = getDisplayImdFromIndex(MI_SHOCK)
            effect_.size = 50
            effect_.velocity.y = 0!
        Case Else
            If Not (effect_.imd <> 0) Then ASSERT_FAILURE "effect.imd != nullptr", "src/effects.cpp:1994", "effectSetupExplosion", "Explosion type unknown!": Exit Sub
        End Select
    End If

    If effect_.typ = EXPLOSION_TYPE_FLAMETHROWER Then
        effect_.frameDelay = 45
    ' Set how long it lasts
    ElseIf effect_.typ = EXPLOSION_TYPE_LASER Then
        effect_.frameDelay = (EXPLOSION_FRAME_DELAY() \ 2) And &HFFFF&
    ElseIf effect_.typ = EXPLOSION_TYPE_TESLA Then
        effect_.frameDelay = EXPLOSION_TESLA_FRAME_DELAY
    ElseIf effect_.typ = EXPLOSION_TYPE_PLASMA Then
        effect_.frameDelay = EXPLOSION_PLASMA_FRAME_DELAY
    ElseIf effect_.typ = EXPLOSION_TYPE_LAND_LIGHT Then
        effect_.frameDelay = 120
    Else
        effect_.frameDelay = EXPLOSION_FRAME_DELAY() And &HFFFF&
    End If


    If effect_.typ = EXPLOSION_TYPE_SHOCKWAVE Then
        effect_.lifeSpan = GAME_TICKS_PER_SEC
    Else
        ' (VB) effectGetNumFrames: imd is null where the misc models are not loaded (C++ would dereference it)
        effect_.lifeSpan = (effect_.frameDelay * effectGetNumFrames(effect_)) And &HFFFF&
    End If


    If (effect_.typ <> EXPLOSION_TYPE_NOT_FACING) And (effect_.typ <> EXPLOSION_TYPE_SHOCKWAVE) Then
        effect_.control = effect_.control Or EFFECT_FACING
    End If
    ' Randomly flip x and y for variation
    If ONEINTWO() Then
        effect_.control = effect_.control Or EFFECT_X_FLIP
    End If
    If ONEINTWO() Then
        effect_.control = effect_.control Or EFFECT_Y_FLIP
    End If
End Sub

Private Sub effectSetupConstruction(effect_ As EFFECT)
    effect_.velocity.x = 0! '(1-rand()%3);
    effect_.velocity.z = 0! '(1-rand()%3);
    effect_.velocity.y = CSng(0 - rand() Mod 3)
    effect_.frameDelay = CONSTRUCTION_FRAME_DELAY() And &HFFFF&
    effect_.imd = getDisplayImdFromIndex(MI_CONSTRUCTION)
    effect_.lifeSpan = CONSTRUCTION_LIFESPAN

    ' These effects always face you
    effect_.control = effect_.control Or EFFECT_FACING

    ' It's a cyclic anim - dies on age
    effect_.control = effect_.control Or EFFECT_CYCLIC

    ' Randomly flip the construction graphics in x and y for variation
    If ONEINTWO() Then
        effect_.control = effect_.control Or EFFECT_X_FLIP
    End If
    If ONEINTWO() Then
        effect_.control = effect_.control Or EFFECT_Y_FLIP
    End If
End Sub

Private Sub effectSetupFire(effect_ As EFFECT)
    effect_.frameDelay = 300    ' needs to be investigated...
    effect_.radius = auxVar And &HFFFF&  ' needs to be investigated
    effect_.lifeSpan = auxVarSec And &HFFFF&
    effect_.birthTime = graphicsTime
    effect_.control = effect_.control Or EFFECT_ESSENTIAL

End Sub

Private Sub effectSetupWayPoint(effect_ As EFFECT)
    If pProximityMsgIMD <> 0 Then effect_.imd = iIMDBaseShape_displayModel(pProximityMsgIMD) Else effect_.imd = 0

    ' These effects musnt make way for others
    effect_.control = effect_.control Or EFFECT_ESSENTIAL
End Sub


Private Sub effectSetupBlood(effect_ As EFFECT)
    effect_.frameDelay = BLOOD_FRAME_DELAY
    effect_.velocity.y = CSng(BLOOD_FALL_SPEED())
    effect_.imd = getDisplayImdFromIndex(MI_BLOOD)
    effect_.size = BLOOD_SIZE() And &HFF&
End Sub

Private Sub effectSetupDestruction(effect_ As EFFECT)
    If effect_.typ = DESTRUCTION_TYPE_SKYSCRAPER Then
        effect_.lifeSpan = ((3 * GAME_TICKS_PER_SEC) \ 2 + (rand() Mod GAME_TICKS_PER_SEC)) And &HFFFF&
        effect_.frameDelay = DESTRUCTION_FRAME_DELAY \ 2
    ElseIf effect_.typ = DESTRUCTION_TYPE_DROID Then
        ' It's all over quickly for droids
        effect_.lifeSpan = DROID_DESTRUCTION_DURATION
        effect_.frameDelay = DESTRUCTION_FRAME_DELAY
    ElseIf effect_.typ = DESTRUCTION_TYPE_WALL_SECTION Or _
        effect_.typ = DESTRUCTION_TYPE_FEATURE Then
        effect_.lifeSpan = STRUCTURE_DESTRUCTION_DURATION \ 4
        effect_.frameDelay = DESTRUCTION_FRAME_DELAY \ 2
    ElseIf effect_.typ = DESTRUCTION_TYPE_POWER_STATION Then
        effect_.lifeSpan = STRUCTURE_DESTRUCTION_DURATION \ 2
        effect_.frameDelay = DESTRUCTION_FRAME_DELAY \ 4
    Else
        ' building's destruction is longer
        effect_.lifeSpan = STRUCTURE_DESTRUCTION_DURATION
        effect_.frameDelay = DESTRUCTION_FRAME_DELAY \ 2
    End If
End Sub


' #define SMOKE_SHIFT (16 - (rand()%32))
Public Sub initPerimeterSmoke(ByVal pImd As Long, base As Vector3i)
    Dim i As Long, shift As Long, inStart As Long, inEnd As Long, varStart As Long, varEnd As Long, varStride As Long
    Dim pos As Vector3i, mnx As Long, mny As Long, mnz As Long, mxx As Long, mxy As Long, mxz As Long
    If vbEffectsOff Then Exit Sub
    shift = 16 - (rand() Mod 32)
    iIMDShape_min pImd, mnx, mny, mnz
    iIMDShape_max pImd, mxx, mxy, mxz
    inStart = mnz - 16
    inEnd = mxz + 16
    varStart = mnx - 16
    varEnd = mxx + 16
    varStride = 24

    For i = varStart To varEnd - 1 Step varStride
        pos = Vector3i_(base.x + i + shift, base.y + 0, base.z + inStart + shift)

        If rand() Mod 6 = 1 Then
            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_SMALL, False, 0, 0, graphicsTime
        Else
            addEffect pos, EFFECT_SMOKE, SMOKE_TYPE_BILLOW, False, 0, 0, graphicsTime
        End If

        pos = Vector3i_(base.x + i + shift, base.y + 0, base.z + inEnd + shift)

        If rand() Mod 6 = 1 Then
            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_SMALL, False, 0, 0, graphicsTime
        Else
            addEffect pos, EFFECT_SMOKE, SMOKE_TYPE_BILLOW, False, 0, 0, graphicsTime
        End If
    Next

    varStart = mnz - 16
    varEnd = mxz + 16
    varStride = 24

    inStart = mnx - 16
    inEnd = mxx + 16

    For i = varStart To varEnd - 1 Step varStride
        pos = Vector3i_(base.x + inStart + shift, base.y + 0, base.z + i + shift)

        If rand() Mod 6 = 1 Then
            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_SMALL, False, 0, 0, graphicsTime
        Else
            addEffect pos, EFFECT_SMOKE, SMOKE_TYPE_BILLOW, False, 0, 0, graphicsTime
        End If

        pos = Vector3i_(base.x + inEnd + shift, base.y + 0, base.z + i + shift)

        If rand() Mod 6 = 1 Then
            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_SMALL, False, 0, 0, graphicsTime
        Else
            addEffect pos, EFFECT_SMOKE, SMOKE_TYPE_BILLOW, False, 0, 0, graphicsTime
        End If
    Next
End Sub


Private Function effectGetNumFrames(psEffect As EFFECT) As Long
    If psEffect.imd <> 0 Then
        effectGetNumFrames = iIMDShape_numFrames(psEffect.imd)
    Else
        effectGetNumFrames = 0
    End If
End Function


Public Sub effectGiveAuxVar(ByVal var As Long)
    auxVar = var
End Sub


Public Sub effectGiveAuxVarSec(ByVal var As Long)
    auxVarSec = var
End Sub

'* Runs all the structure effect stuff - steam puffing out etc
Private Sub effectStructureUpdates()
    Dim frameNumber As Long, curPartition As Long, player As Long, it As Long, psStructure As Long
    Dim partition As Long, pDisplayModel As Long, numConnectors As Long, orderAdjustment As Long
    Dim idx As Long, orderIdx As Long, effectStartTime As Long, eventPos As Vector3i, af As Affine3F
    Dim c As Vector3i, cx As Long, cy As Long, cz As Long, rotated As Vector3i

    frameNumber = frameGetFrameNumber()
    curPartition = FromU32(U32(frameNumber) - Int(U32(frameNumber) / EFFECT_STRUCTURE_DIVISION) * EFFECT_STRUCTURE_DIVISION)
    ' Is it the right time?
    If Not ULT(lastUpdateStructures(curPartition) + STRUCTURE_UPDATE_INTERVAL, graphicsTime) Then
        Exit Sub
    End If
    ' Store away the last update time.
    lastUpdateStructures(curPartition) = graphicsTime

    ' Go thru' all players
    For player = 0 To MAX_PLAYERS - 1
        it = apsStructLists(player).head
        Do While it <> 0
            psStructure = lnObj(it)
            ' Find its group.
            partition = FromU32(U32(aStruct(psStructure).id) - Int(U32(aStruct(psStructure).id) / EFFECT_STRUCTURE_DIVISION) * EFFECT_STRUCTURE_DIVISION)

            ' Is it the right frame?
            If partition <> curPartition Then
                GoTo continue_
            End If

            If aStruct(psStructure).status <> SS_BUILT Or BASE_OBJECT_visibleForLocalDisplay(psStructure) = 0 Then
                GoTo continue_
            End If

            pDisplayModel = iIMDBaseShape_displayModel(aStruct(psStructure).sDisplay.imd)

            ' Factories puff out smoke, power stations puff out tesla stuff
            Select Case asStructureStats(aStruct(psStructure).pStructureType).typ
            Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
                '
                '   We're a factory, so better puff out a bit of steam
                '   Complete hack with the magic numbers - just for IAN demo
                '
                ' Display only - use the display model
                If iIMDShape_connectors_size(pDisplayModel) > 0 Then
                    numConnectors = iIMDShape_connectors_size(pDisplayModel)
                    orderAdjustment = FromU32(U32(frameNumber) - Int(U32(frameNumber) / numConnectors) * numConnectors)
                    For idx = 0 To numConnectors - 1
                        iIMDShape_connector pDisplayModel, idx, cx, cy, cz
                        Affine3F_construct af
                        Affine3F_RotY af, aStruct(psStructure).rot.direction
                        c = Vector3i_(cx, cz, -cy)
                        rotated = Affine3F_mul_vec(af, c)
                        ' psStructure->pos.xzy() + ...
                        eventPos.x = aStruct(psStructure).pos.x + rotated.x
                        eventPos.y = aStruct(psStructure).pos.z + rotated.y
                        eventPos.z = aStruct(psStructure).pos.y + rotated.z

                        orderIdx = idx + orderAdjustment
                        If orderIdx >= numConnectors Then
                            orderIdx = orderIdx - numConnectors
                        End If
                        effectStartTime = graphicsTime + 150 * orderIdx
                        addEffect eventPos, EFFECT_SMOKE, SMOKE_TYPE_STEAM, False, 0, 0, effectStartTime
                    Next
                    audio_PlayObjStaticTrack psStructure, ID_SOUND_STEAM
                End If
            Case REF_POWER_GEN
                eventPos.x = aStruct(psStructure).pos.x
                eventPos.y = aStruct(psStructure).pos.z
                eventPos.z = aStruct(psStructure).pos.y

                ' Add an effect over the central spire.

                eventPos.y = aStruct(psStructure).pos.z + 48

                addEffect eventPos, EFFECT_EXPLOSION, EXPLOSION_TYPE_TESLA, False, 0, 0, graphicsTime

                audio_PlayObjStaticTrack psStructure, ID_SOUND_POWER_SPARK
                audio_PlayObjStaticTrack psStructure, ID_SOUND_POWER_HUM
            Case REF_RESOURCE_EXTRACTOR
                If aStruct(psStructure).pFunctionality <> 0 Then
                    If aStruct(psStructure).functionality.resourceExtractor.psPowerGen <> 0 And _
                       aStruct(psStructure).animationEvent = ANIM_EVENT_ACTIVE Then
                        audio_PlayObjStaticTrack psStructure, ID_SOUND_OIL_PUMP_2
                    End If
                End If
            Case Else
            End Select
continue_:
            it = lnNext(it)
        Loop
    Next
End Sub

Public Sub effectResetUpdates()
    Dim i As Long
    For i = 0 To EFFECT_STRUCTURE_DIVISION - 1
        lastUpdateStructures(i) = 0
    Next
End Sub

' glm vec3 / ivec3 to_json (wzjsonhelpers): an array of the three numbers
Private Function vec3f_to_json(v As Vector3f) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_array()
    j.push_back json_value(CDbl(v.x))
    j.push_back json_value(CDbl(v.y))
    j.push_back json_value(CDbl(v.z))
    Set vec3f_to_json = j
End Function

Private Function vec3i_to_json(v As Vector3i) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_array()
    j.push_back json_value(v.x)
    j.push_back json_value(v.y)
    j.push_back json_value(v.z)
    Set vec3i_to_json = j
End Function

'* This will save out the effects data
Public Function writeFXData(ByVal fileName As String) As Boolean
    Dim i As Long, it As Long, mRoot As nlohmann_json, effectObj As nlohmann_json, effectKey As String
    Dim jsonString As String, bytes() As Byte, n As Long
    i = 0
    Set mRoot = json_object()
    For it = 1 To effectUsed
        If effectAlive(it) Then
            Set effectObj = json_object()
            effectObj.set_item "control", json_value_unsigned(aEffect(it).control)
            effectObj.set_item "group", json_value(CLng(aEffect(it).group))
            effectObj.set_item "type", json_value(CLng(aEffect(it).typ))
            effectObj.set_item "frameNumber", json_value_unsigned(aEffect(it).frameNumber)
            effectObj.set_item "size", json_value_unsigned(aEffect(it).size)
            effectObj.set_item "baseScale", json_value_unsigned(aEffect(it).baseScale)
            effectObj.set_item "specific", json_value_unsigned(aEffect(it).specific)
            effectObj.set_item "position", vec3f_to_json(aEffect(it).position)
            effectObj.set_item "velocity", vec3f_to_json(aEffect(it).velocity)
            effectObj.set_item "rotation", vec3i_to_json(aEffect(it).rotation)
            effectObj.set_item "spin", vec3i_to_json(aEffect(it).spin)
            effectObj.set_item "birthTime", json_value_unsigned(U32(aEffect(it).birthTime))
            effectObj.set_item "lastFrame", json_value_unsigned(U32(aEffect(it).lastFrame))
            effectObj.set_item "frameDelay", json_value_unsigned(aEffect(it).frameDelay)
            effectObj.set_item "lifeSpan", json_value_unsigned(aEffect(it).lifeSpan)
            effectObj.set_item "radius", json_value_unsigned(aEffect(it).radius)

            If aEffect(it).imd <> 0 Then
                effectObj.set_item "imd_name", json_value(modelName(aEffect(it).imd))
            End If

            effectKey = "effect" & CStr(i)
            mRoot.set_item effectKey, effectObj

            ' Move on to reading the next effect
            i = i + 1
        End If
    Next

    On Error GoTo dump_failed
    jsonString = mRoot.dump(4)
    On Error GoTo 0
    debug_ 2380, LOG_SAVE, "writeFXData", "%s %s", "Saving", fileName
    n = utf8_encode(jsonString, bytes)
    saveFile fileName, bytes, n

    ' Everything is just fine!
    writeFXData = True
    Exit Function
dump_failed:
    ASSERT_FAILURE "false", "src/effects.cpp:2377", "writeFXData", "Failed to save JSON to %s with error: %s", fileName, Err.Description
    writeFXData = False
End Function

'* This will read in the effects data
Public Function readFXData(ByVal fileName As String) As Boolean
    Dim ini As WzConfig, list() As String, n As Long, i As Long, curEffect As EFFECT
    Dim imd_name As String, baseImd As Long, timeThatEffectHasBeenRunning As Long, timeLeftToRun As Long
    ' Clear out anything that's there already!
    initEffectsSystem

    Set ini = WzConfig_new(fileName, WzConfig_ReadOnly)
    n = ini.childGroups(list)

    For i = 0 To n - 1
        ini.beginGroup list(i)
        curEffect = EFFECT_new()

        curEffect.control = ini.value("control").ToInt() And &HFF&
        curEffect.group = ini.value("group").ToInt()
        curEffect.typ = ini.value("type").ToInt()
        curEffect.frameNumber = ini.value("frameNumber").ToInt() And &HFF&
        curEffect.size = ini.value("size").ToInt() And &HFFFF&
        curEffect.baseScale = ini.value("baseScale").ToInt() And &HFF&
        curEffect.specific = ini.value("specific").ToInt() And &HFF&
        ini.vector3f "position", curEffect.position.x, curEffect.position.y, curEffect.position.z
        ini.vector3f "velocity", curEffect.velocity.x, curEffect.velocity.y, curEffect.velocity.z
        ini.vector3i "rotation", curEffect.rotation.x, curEffect.rotation.y, curEffect.rotation.z
        ini.vector3i "spin", curEffect.spin.x, curEffect.spin.y, curEffect.spin.z
        curEffect.birthTime = ini.value("birthTime").ToInt()
        curEffect.lastFrame = ini.value("lastFrame").ToInt()
        curEffect.frameDelay = ini.value("frameDelay").ToInt() And &HFFFF&
        curEffect.lifeSpan = ini.value("lifeSpan").ToInt() And &HFFFF& ' this is the original duration of the efect, not the time remaining
        curEffect.radius = ini.value("radius").ToInt() And &HFFFF&
        If ini.contains("imd_name") Then
            imd_name = ini.value("imd_name").toWzString()
            If Len(imd_name) <> 0 Then
                baseImd = modelGet(imd_name)
                If baseImd <> 0 Then curEffect.imd = iIMDBaseShape_displayModel(baseImd) Else curEffect.imd = 0
            End If
        Else
            curEffect.imd = 0
        End If

        ' For fire effects, set the tile as being on fire so that (e.g.) burning oil resources can't
        ' immediately be built on
        If EFFECT_FIRE = curEffect.group Then
            timeThatEffectHasBeenRunning = curEffect.lastFrame - curEffect.birthTime
            timeLeftToRun = curEffect.lifeSpan - timeThatEffectHasBeenRunning

            ' Sanity check - don't allow a negative time to wrap to a huge positive unsigned value.
            If timeLeftToRun > 0 Then
                tileSetFire ToInt(curEffect.position.x), ToInt(curEffect.position.z), timeLeftToRun
            End If
        End If

        ' Move on to reading the next effect
        ini.endGroup

        gActiveEffects_emplace curEffect
    Next

    ' Hopefully everything's just fine by now
    readFXData = True
End Function
