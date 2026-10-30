Attribute VB_Name = "atmos_c"
' src/atmos.cpp, atmos.h: the weather (rain and snow particles), whole.
' asAtmosParts is a dynamic array (asAtmosPartsAllocated: != nullptr); an
' ATPART * (renderParticle's, bucket3d's RENDER_PARTICLE) is its index + 1
' (0 = nullptr). The particles' display randomness is cstdlib_vb's rand.
' gameTimeGetMod().asDouble() is gtime_c's numerator / denominator (taken as
' NaN when the denominator is 0: then nothing is added). weather's initialiser
' (WT_NONE, not VB's 0 = WT_RAINING) is set on the first call (atmosInitVars).
Option Explicit

Public Type ATPART
    status As Long                      ' UBYTE
    typ As Long                         ' UBYTE
    size As Long                        ' UDWORD
    position As Vector3f
    velocity As Vector3f
    imd As Long                         ' iIMDBaseShape *
End Type

Public Enum WT_CLASS
    WT_RAINING
    WT_SNOWING
    WT_NONE
End Enum

' -----------------------------------------------------------------------------
' -----------------------------------------------------------------------------
' Shift all this gubbins into a .h file if it makes it into game
' -----------------------------------------------------------------------------
' Roughly one per tile
Private Const MAX_ATMOS_PARTICLES As Long = 65536    ' (MAP_MAXWIDTH * MAP_MAXHEIGHT)

Private Enum AP_TYPE
    AP_RAIN
    AP_SNOW
End Enum

Private Enum AP_STATUS
    APS_INACTIVE
    APS_ACTIVE
End Enum

Public asAtmosParts() As ATPART         ' = nullptr (0-based)
Private asAtmosPartsAllocated As Boolean
Private freeParticle As Long            ' UDWORD (= 0)
Private weather As WT_CLASS             ' = WT_NONE (atmosInitVars)
Private atmosVarsSet As Boolean
Private accumulatedParticlesToAdd As Double ' atmosUpdateSystem's static (= 0.0)

' (VB) the globals' C++ initialisers
Private Sub atmosInitVars()
    If atmosVarsSet Then Exit Sub
    atmosVarsSet = True
    weather = WT_NONE
End Sub

Private Function SNOW_SPEED_DRIFT() As Long
    SNOW_SPEED_DRIFT = 40 - rand() Mod 80
End Function

Private Function SNOW_SPEED_FALL() As Long
    SNOW_SPEED_FALL = 0 - (rand() Mod 40 + 80)
End Function

Private Function RAIN_SPEED_DRIFT() As Long
    RAIN_SPEED_DRIFT = rand() Mod 50
End Function

Private Function RAIN_SPEED_FALL() As Long
    RAIN_SPEED_FALL = 0 - ((rand() Mod 300) + 700)
End Function

' Setup all the particles
Public Sub atmosInitSystem()
    atmosInitVars
    If Not asAtmosPartsAllocated And weather <> WT_NONE Then
        ' calloc sets all to APS_INACTIVE initially
        ReDim asAtmosParts(0 To MAX_ATMOS_PARTICLES - 1)
        asAtmosPartsAllocated = True
    End If
    ' Start at the beginning
    freeParticle = 0
End Sub

'   Makes a particle wrap around - if it goes off the grid, then it returns
'   on the other side - provided it's still on world... Which it should be
Private Sub testParticleWrap(psPart As ATPART)
    ' Gone off left side
    If psPart.position.x < playerPos.p.x - world_coord(visibleTiles.x) \ 2 Then
        psPart.position.x = psPart.position.x + world_coord(visibleTiles.x)

    ' Gone off right side
    ElseIf psPart.position.x > (playerPos.p.x + world_coord(visibleTiles.x) \ 2) Then
        psPart.position.x = psPart.position.x - world_coord(visibleTiles.x)
    End If

    ' Gone off top
    If psPart.position.z < playerPos.p.z - world_coord(visibleTiles.y) \ 2 Then
        psPart.position.z = psPart.position.z + world_coord(visibleTiles.y)

    ' Gone off bottom
    ElseIf psPart.position.z > (playerPos.p.z + world_coord(visibleTiles.y) \ 2) Then
        psPart.position.z = psPart.position.z - world_coord(visibleTiles.y)
    End If
End Sub

' Moves one of the particles
Private Sub processParticle(psPart As ATPART)
    Dim groundHeight As Long, pos As Vector3i, x As Long, y As Long, psTile As Long

    ' Only move if the game isn't paused
    If Not gamePaused() Then
        ' Move the particle - frame rate controlled
        psPart.position.x = psPart.position.x + graphicsTimeAdjustedIncrement(psPart.velocity.x)
        psPart.position.y = psPart.position.y + graphicsTimeAdjustedIncrement(psPart.velocity.y)
        psPart.position.z = psPart.position.z + graphicsTimeAdjustedIncrement(psPart.velocity.z)

        ' Wrap it around if it's gone off grid...
        testParticleWrap psPart

        ' If it's gone off the WORLD...
        If psPart.position.x < 0 Or psPart.position.z < 0 Or _
           psPart.position.x > ((mapWidth - 1) * TILE_UNITS) Or _
           psPart.position.z > ((mapHeight - 1) * TILE_UNITS) Then
            ' The kill it
            psPart.status = APS_INACTIVE
            Exit Sub
        End If

        ' What height is the ground under it? Only do if low enough...
        If psPart.position.y < TILE_MAX_HEIGHT Then
            ' Get ground height
            groundHeight = map_Height(Fix(psPart.position.x), Fix(psPart.position.z))

            ' Are we below ground?
            If Fix(psPart.position.y) < groundHeight _
               Or psPart.position.y < 0! Then
                ' Kill it and return
                psPart.status = APS_INACTIVE
                If psPart.typ = AP_RAIN Then
                    x = map_coord(Fix(psPart.position.x))
                    y = map_coord(Fix(psPart.position.z))
                    psTile = mapTile(x, y)
                    If terrainType(psTile) = TER_WATER Then
                        If TEST_TILE_VISIBLE_TO_SELECTEDPLAYER(psTile) Then ' display-only check for adding effect
                            pos.x = Fix(psPart.position.x)
                            pos.z = Fix(psPart.position.z)
                            pos.y = groundHeight
                            effectSetSize 60
                            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_SPECIFIED, True, getDisplayImdFromIndex(MI_SPLASH), 0, graphicsTime
                        End If
                    End If
                End If
                Exit Sub
            End If
        End If
        If psPart.typ = AP_SNOW Then
            If rand() Mod 30 = 1 Then
                psPart.velocity.z = CSng(SNOW_SPEED_DRIFT())
            End If
            If rand() Mod 30 = 1 Then
                psPart.velocity.x = CSng(SNOW_SPEED_DRIFT())
            End If
        End If
    End If
End Sub

' Adds a particle to the system if it can
Private Sub atmosAddParticle(pos As Vector3f, ByVal typ As AP_TYPE)
    Dim activeCount As Long, i As Long

    i = freeParticle
    activeCount = 0
    Do While asAtmosParts(i).status = APS_ACTIVE And activeCount < MAX_ATMOS_PARTICLES
        activeCount = activeCount + 1
        ' Check for wrap around
        If i >= (MAX_ATMOS_PARTICLES - 1) Then
            ' Go back to the first one
            i = 0
        End If
        i = i + 1
    Loop

    ' Check the list isn't just full of essential effects
    If activeCount >= MAX_ATMOS_PARTICLES - 1 Then
        ' All of the particles active!?!?
        Exit Sub
    Else
        freeParticle = i
    End If

    ' Record it's type
    asAtmosParts(freeParticle).typ = typ And &HFF&

    ' Make it active
    asAtmosParts(freeParticle).status = APS_ACTIVE

    ' Setup the imd
    Select Case typ
    Case AP_SNOW
        asAtmosParts(freeParticle).imd = getImdFromIndex(MI_SNOW)
        asAtmosParts(freeParticle).size = 80
    Case AP_RAIN
        asAtmosParts(freeParticle).imd = getImdFromIndex(MI_RAIN)
        asAtmosParts(freeParticle).size = 50
    Case Else
    End Select

    ' Setup position
    asAtmosParts(freeParticle).position = pos

    ' Setup its velocity (the Vector3f's arguments taken left to right)
    If typ = AP_RAIN Then
        asAtmosParts(freeParticle).velocity.x = RAIN_SPEED_DRIFT()
        asAtmosParts(freeParticle).velocity.y = RAIN_SPEED_FALL()
        asAtmosParts(freeParticle).velocity.z = RAIN_SPEED_DRIFT()
    Else
        asAtmosParts(freeParticle).velocity.x = SNOW_SPEED_DRIFT()
        asAtmosParts(freeParticle).velocity.y = SNOW_SPEED_FALL()
        asAtmosParts(freeParticle).velocity.z = SNOW_SPEED_DRIFT()
    End If
End Sub

' Move the particles
Public Sub atmosUpdateSystem()
    Dim i As Long, numberToAdd As Long, pos As Vector3f, p As ATPART
    Dim modNumerator As Long, modDenominator As Long, gameTimeModVal As Double

    atmosInitVars
    If Not asAtmosPartsAllocated Then
        Exit Sub
    End If

    ' we don't want to do any of this while paused.
    If Not gamePaused() And weather <> WT_NONE Then
        For i = 0 To MAX_ATMOS_PARTICLES - 1
            ' See if it's active
            If asAtmosParts(i).status = APS_ACTIVE Then
                ' (VB) in a copy: no element passed ByRef while the update adds effects
                p = asAtmosParts(i)
                processParticle p
                asAtmosParts(i) = p
            End If
        Next

        ' The original code added a fixed number of particles per tick. To take into account game speed
        ' we have to accumulate a fractional number of particles to add them at a slower or faster rate.
        gameTimeGetMod modNumerator, modDenominator
        If modDenominator <> 0 Then
            gameTimeModVal = CDbl(modNumerator) / CDbl(modDenominator)
            If weather = WT_SNOWING Then
                accumulatedParticlesToAdd = accumulatedParticlesToAdd + 2# * gameTimeModVal
            Else
                accumulatedParticlesToAdd = accumulatedParticlesToAdd + 4# * gameTimeModVal
            End If
        End If

        numberToAdd = Fix(accumulatedParticlesToAdd)
        accumulatedParticlesToAdd = accumulatedParticlesToAdd - numberToAdd

        ' Temporary stuff - just adds a few particles!
        For i = 0 To numberToAdd - 1
            pos.x = playerPos.p.x
            pos.z = playerPos.p.z
            pos.x = pos.x + world_coord(rand() Mod visibleTiles.x - visibleTiles.x \ 2)
            pos.z = pos.z + world_coord(rand() Mod visibleTiles.x - visibleTiles.y \ 2)
            pos.y = 1000

            ' If we've got one on the grid
            If pos.x > 0 And pos.z > 0 And _
               pos.x < world_coord(mapWidth - 1) And _
               pos.z < world_coord(mapHeight - 1) Then
                ' On grid, so which particle shall we add?
                Select Case weather
                Case WT_SNOWING
                    atmosAddParticle pos, AP_SNOW
                Case WT_RAINING
                    atmosAddParticle pos, AP_RAIN
                Case WT_NONE
                End Select
            End If
        Next
    End If
End Sub

Private Sub renderParticleInternal(psPart As ATPART, viewMatrix As mat4, rotateScaleMatrix As mat4)
    Dim dv As Vector3f, modelMatrix As mat4

    ' Transform it
    dv.x = psPart.position.x
    dv.y = psPart.position.y
    dv.z = -(psPart.position.z)
    ' Make it face camera
    ' Scale it...
    modelMatrix = glm_mat4_mul(glm_translate(dv), rotateScaleMatrix)
    pie_Draw3DShape iIMDBaseShape_displayModel(psPart.imd), 0, 0, psPalette(WZCOL_WHITE), 0, 0, VarPtr(modelMatrix), VarPtr(viewMatrix)
    ' Draw it...
End Sub

Public Sub atmosDrawParticles(viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim i As Long, rotateMatrix As mat4, rotateScaleMatrix As mat4, last_particle_size As Long, s As Single

    atmosInitVars
    If weather = WT_NONE Or Not asAtmosPartsAllocated Then
        Exit Sub
    End If

    rotateMatrix = glm_mat4_mul(glm_rotate(UNDEG(-playerPos.r.y), glm_vec3(0!, 1!, 0!)), _
                                glm_rotate(UNDEG(-playerPos.r.x), glm_vec3(0!, 1!, 0!)))
    rotateScaleMatrix = rotateMatrix
    last_particle_size = 0

    ' Traverse the list
    For i = 0 To MAX_ATMOS_PARTICLES - 1
        ' Don't bother unless it's active
        If asAtmosParts(i).status = APS_ACTIVE Then
            ' Is it visible on the screen?
            If clipXYZ(Fix(asAtmosParts(i).position.x), Fix(asAtmosParts(i).position.z), Fix(asAtmosParts(i).position.y), perspectiveViewMatrix) Then
                If last_particle_size <> asAtmosParts(i).size Then
                    s = asAtmosParts(i).size / 100!
                    rotateScaleMatrix = glm_mat4_mul(rotateMatrix, glm_scale(glm_vec3(s, s, s)))
                    last_particle_size = asAtmosParts(i).size
                End If
                renderParticleInternal asAtmosParts(i), viewMatrix, rotateScaleMatrix
            End If
        End If
    Next
End Sub

' psPart: an ATPART * (its index + 1)
Public Sub renderParticle(ByVal psPart As Long, viewMatrix As mat4)
    Dim rotateScaleMatrix As mat4, s As Single
    s = asAtmosParts(psPart - 1).size / 100!
    rotateScaleMatrix = glm_mat4_mul(glm_mat4_mul(glm_rotate(UNDEG(-playerPos.r.y), glm_vec3(0!, 1!, 0!)), _
        glm_rotate(UNDEG(-playerPos.r.x), glm_vec3(0!, 1!, 0!))), _
        glm_scale(glm_vec3(s, s, s)))
    renderParticleInternal asAtmosParts(psPart - 1), viewMatrix, rotateScaleMatrix
End Sub

Public Sub atmosSetWeatherType(ByVal typ As WT_CLASS)
    atmosInitVars
    If typ <> weather Then
        weather = typ
        atmosInitSystem
    End If
    If typ = WT_NONE And asAtmosPartsAllocated Then
        Erase asAtmosParts
        asAtmosPartsAllocated = False
    End If
End Sub

Public Function atmosGetWeatherType() As WT_CLASS
    atmosInitVars
    atmosGetWeatherType = weather
End Function
