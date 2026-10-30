Attribute VB_Name = "display3d_c"
' src/display3d.cpp (part: lines 1-399, the globals; 584-2498, the scene, the
' clipping, the objects' display lists, the blueprints, the view position;
' 3925-4102, the screen coordinates, the mouse tile, the sky and the camera
' height), display3d.h, display3ddef.h. The other parts are
' display3d_render_c (2499-3167: renderFeature ... renderWallSection) and
' display3d_overlay_c (400-583: BatchedObjectStatusRenderer; 3168-3924: the 2D
' overlays; 4103-4637: targets, structure effects, ranges, rank graphics,
' construction lines). display3d.cpp's file-level statics and globals are all
' here, Public where another part uses them.
' A glm::mat4 is glm_h's mat4 (ByRef as C++'s const &); WzGfx takes it by
' address (VarPtr). The drawing goes through WzGfx (GfxApi): pie_*, the
' context's gfx_api_context_*, gfx_api_pso_get. A display shape (iIMDShape *)
' and a base shape (iIMDBaseShape *) are WzGfx's handles (0 = nullptr), read
' with GfxApi's iIMDShape_<field> / iIMDBaseShape_<field>.
' A Blueprint's optional<STRUCTURE> is a structure handle (0 = nullopt): see
' buildBlueprint. QUAD is geometry_c's. The widgets of the network status
' form (setupConnectionStatusForm, NetworkDisplayImage) wait for cluster G.
Option Explicit

' display3ddef.h
Public Const TILE_WIDTH As Long = 128
Public Const TILE_HEIGHT As Long = 128
Public Const TILE_SIZE As Long = TILE_WIDTH * TILE_HEIGHT
Public Const VISIBLE_XTILES As Long = 64
Public Const VISIBLE_YTILES As Long = 64
Public Const SKY_MULT As Long = 1

' display3d.h
Public Const HEIGHT_TRACK_INCREMENTS As Long = 50

' enum TILE_ID
Public Const RIVERBED_TILE As Long = 5          ' Underwater ground
Public Const WATER_TILE As Long = 17            ' Water surface
Public Const RUBBLE_TILE As Long = 54           ' You can drive over these
Public Const BLOCKING_RUBBLE_TILE As Long = 67  ' You cannot drive over these

Public Enum ENERGY_BAR
    BAR_SELECTED
    BAR_DROIDS
    BAR_DROIDS_AND_STRUCTURES
    BAR_LAST
End Enum

Public Type iView
    p As Vector3i                       ' = Vector3i(0, 0, 0)
    r As Vector3i                       ' = Vector3i(0, 0, 0)
End Type

Public Const BASE_MUZZLE_FLASH_DURATION As Long = GAME_TICKS_PER_SEC \ 10
Public Const EFFECT_MUZZLE_ADDITIVE As Long = 128

' (VB) ivisdef.h's ANIMFRAME, as WzGfx's iIMDShape_objanimdata_p copies it (the same layout)
Private Type ANIMFRAME
    scale As Vector3f
    pos As Vector3f
    rot As Rotation
End Type

' lib/wzmaplib/include/wzmaplib/map.h
Private Const ELEVATION_SCALE As Long = 2
' hci.h (cluster G): RET_X; E_H is (pie_GetVideoBufferHeight() - BASE_COORDS_Y)

' imd.h's flags of a display shape (iIMDShape::flags)
Private Const iV_IMD_NO_ADDITIVE As Long = &H1&
Private Const iV_IMD_ADDITIVE As Long = &H2&
Private Const iV_IMD_PREMULTIPLIED As Long = &H4&
Private Const iV_IMD_ROLL_TO_CAMERA As Long = &H10&
Private Const iV_IMD_PITCH_TO_CAMERA As Long = &H20&

Private txtLevelName As WzText
Private txtDebugStatus As WzText
Private txtCurrentTime As WzText
Private txtShowFPS As WzText
Private txtUnits As WzText
' show Samples text
Private txtShowSamples_Que As WzText
Private txtShowSamples_Lst As WzText
Private txtShowSamples_Act As WzText
' show Orders text
Private txtShowOrders As WzText
' show Droid visible/draw counts text
Private droidText As WzText

Private pScreenTriangleVBO As gfx_api_buffer   ' = nullptr

'********************  Variables  ********************
' Should be cleaned up properly and be put in structures.

' Initialised at start of drawTiles().
' In model coordinates where x is east, y is up and z is north, rather than world coordinates where x is east, y is south and z is up.
' To get the real camera position, still need to add Vector3i(player.p.x, 0, player.p.z).
Private actualCameraPosition As Vector3i

Public bRangeDisplay As Boolean         ' = false
Public rangeCenterX As Long, rangeCenterY As Long, rangeRadius As Long
Private bDrawProximitys As Boolean      ' = true
Public godMode As Boolean
Public showGateways As Boolean          ' = false
Public showPath As Boolean              ' = false

' Skybox data
Private wind As Single                  ' = 0.0f
Private windSpeed As Single             ' = 0.0f
Private skybox_scale As Single          ' = 10000.0f

'
Public bDrawTerrainShadows As Boolean   ' = true

' When to display HP bars
Public barMode As Long                  ' UWORD

' Have we made a selection by clicking the mouse? - used for dragging etc
Public selectAttempt As Boolean         ' = false

' Vectors that hold the player and camera directions and positions
Public playerPos As iView

' How far away are we from the terrain
Private distance As Single
Private bIsCloseDistance As Boolean     ' = false

' Stores the screen coordinates of the transformed terrain tiles
Private tileScreenInfo(0 To VISIBLE_YTILES, 0 To VISIBLE_XTILES) As Vector3i
Private tileScreenVisible(0 To VISIBLE_YTILES, 0 To VISIBLE_XTILES) As Boolean   ' = {false}

' Records the present X and Y values for the current mouse tile (in tiles)
Public mouseTileX As Long, mouseTileY As Long
Public mousePos As Vector2i             ' (0, 0)

' Do we want the radar to be rendered
Public radarOnScreen As Boolean         ' = true
Public radarPermitted As Boolean        ' = true

' Show unit/building gun/sensor range
Public rangeOnScreen As Boolean         ' = false: For now, most likely will change later!  -Q 5-10-05   A very nice effect - Per

' Tactical UI: show/hide target origin icon
Public tuiTargetOrigin As Boolean       ' = false

' Temporary values for the terrain render - center of grid to be rendered
Private playerXTile As Long, playerZTile As Long

' The cached value of frameGetFrameNumber()
Public currentGameFrame As Long         ' UDWORD
' The box used for multiple selection - present screen coordinates
Private dragQuad As QUAD

'* Number of tiles visible
' * \todo This should become dynamic! (A function of resolution, angle and zoom maybe.)
Public visibleTiles As Vector2i         ' const Vector2i visibleTiles(VISIBLE_XTILES, VISIBLE_YTILES)

' The tile we use for drawing the bottom of a body of water
Public underwaterTile As Long           ' = WATER_TILE
'* The tile we use for drawing rubble
' * \note Unused.
Public rubbleTile As Long               ' = BLOCKING_RUBBLE_TILE

'* Show how many frames we are rendering per second
' * default OFF, turn ON via console command 'showfps'
Public showFPS As Boolean               ' = false
'* Show how many samples we are rendering per second
' * default OFF, turn ON via console command 'showsamples'
Public showUNITCOUNT As Boolean         ' = false
'* Show how many kills/deaths (produced units) made
' * default OFF, turn ON via console command 'showunits'
Public showSAMPLES As Boolean           ' = false
'*  Show the current selected units order / action
' *  default OFF, turn ON via console command 'showorders'
Public showORDERS As Boolean            ' = false
'*  Show the drawn/undrawn counts for droids
'  * default OFF, turn ON by flipping it here
Public showDROIDcounts As Boolean       ' = false

'*  Speed of blueprints animation (moving from one tile to another)
'  * default 20, change in config
Public BlueprintTrackAnimationSpeed As Long ' = 20

'* When we have a connection issue, we will flash a message on screen
Private errorWaiting As String          ' const char * (vbNullString = nullptr)
Private lastErrorTime As Long           ' uint32_t = 0

Private Const NETWORK_FORM_ID As Long = &HFAAA&
Private Const NETWORK_BUT_ID As Long = &HFAAB&
'* When enabled, this causes a segfault in the game, to test out the crash handler
Public CauseCrash As Boolean            ' = false

'* tells us in realtime, what droid is doing (order / action)
Public DROIDDOING As String             ' char[512]

' Geometric offset which will be passed to pie_SetGeometricOffset
Private Const geoOffset As Long = 192

' The average terrain height for the center of the area the camera is looking at
Private averageCentreTerrainHeight As Long

'* The time at which a sensor target was last assigned
' * Used to draw a visual effect.
Public lastTargetAssignation As Long    ' UDWORD = 0
'* The time at which an order concerning a destination was last given
' * Used to draw a visual effect.
Public lastDestAssignation As Long      ' UDWORD = 0
Public bSensorTargetting As Boolean     ' = false
Public bDestTargetting As Boolean       ' = false
Public psSensorObj As Long              ' BASE_OBJECT * = nullptr
Public destTargetX As Long, destTargetY As Long     ' UDWORD
Public destTileX As Long, destTileY As Long         ' UDWORD = 0

Public Type Blueprint
    stats As Long                       ' STRUCTURE_STATS const * (-1 = nullptr)
    pos As Vector3i
    dir_ As Long                         ' uint16_t
    index As Long                       ' uint32_t
    state As STRUCT_STATES
    player As Long                      ' uint8_t
End Type

Private blueprints() As Blueprint       ' static std::vector<Blueprint> blueprints;
Private blueprints_size As Long

Public Const TARGET_TO_SENSOR_TIME As Long = (4 * (GAME_TICKS_PER_SEC)) \ 5
Public Const DEST_TARGET_TIME As Long = GAME_TICKS_PER_SEC \ 4

' The distance the selection box will pulse
Public Const BOX_PULSE_SIZE As Single = 30!

' the opacity at which building blueprints will be drawn
Public Const BLUEPRINT_OPACITY As Long = 120

' std::unique_ptr<Blueprint> playerBlueprint = std::make_unique<Blueprint>(); and the ValueTrackers
Private playerBlueprint As Blueprint
Private playerBlueprintX As ValueTracker
Private playerBlueprintY As ValueTracker
Private playerBlueprintZ As ValueTracker
Private playerBlueprintDirection As ValueTracker

' trackHeight's statics
Private trackHeight_lastHeightAdjustmentRealTime As Long    ' uint32_t = 0
Private trackHeight_heightSpeed As Single                   ' = 0.0f

' (VB) the C++ initialisers of the globals above (Sub Main runs it with the other modules' inits)
Public Sub display3d_c_init()
    Set txtLevelName = New WzText
    Set txtDebugStatus = New WzText
    Set txtCurrentTime = New WzText
    Set txtShowFPS = New WzText
    Set txtUnits = New WzText
    Set txtShowSamples_Que = New WzText
    Set txtShowSamples_Lst = New WzText
    Set txtShowSamples_Act = New WzText
    Set txtShowOrders = New WzText
    Set droidText = New WzText
    bDrawProximitys = True
    wind = 0!
    windSpeed = 0!
    skybox_scale = 10000!
    bDrawTerrainShadows = True
    radarOnScreen = True
    radarPermitted = True
    visibleTiles.x = VISIBLE_XTILES
    visibleTiles.y = VISIBLE_YTILES
    underwaterTile = WATER_TILE
    rubbleTile = BLOCKING_RUBBLE_TILE
    BlueprintTrackAnimationSpeed = 20
    errorWaiting = vbNullString
    playerBlueprint = Blueprint_new()
End Sub

' Blueprint() : stats(nullptr), pos({0,0,0}), dir(0), index(0), state(SS_BLUEPRINT_INVALID), player(selectedPlayer)
Private Function Blueprint_new() As Blueprint
    Blueprint_new.stats = -1
    Blueprint_new.state = SS_BLUEPRINT_INVALID
    Blueprint_new.player = selectedPlayer
End Function

' Blueprint(STRUCTURE_STATS const *stats, Vector3i pos, uint16_t dir, uint32_t index, STRUCT_STATES state, uint8_t player)
Private Function Blueprint_(ByVal stats As Long, pos As Vector3i, ByVal dir_ As Long, ByVal index As Long, ByVal state As STRUCT_STATES, ByVal player As Long) As Blueprint
    Blueprint_.stats = stats
    Blueprint_.pos = pos
    Blueprint_.dir_ = dir_ And &HFFFF&
    Blueprint_.index = index
    Blueprint_.state = state
    Blueprint_.player = player And &HFF&
End Function

Private Function Blueprint_compare(a As Blueprint, b As Blueprint) As Long
    If asStructureStats(a.stats).ref <> asStructureStats(b.stats).ref Then
        If asStructureStats(a.stats).ref < asStructureStats(b.stats).ref Then Blueprint_compare = -1 Else Blueprint_compare = 1
        Exit Function
    End If
    If a.pos.x <> b.pos.x Then
        If a.pos.x < b.pos.x Then Blueprint_compare = -1 Else Blueprint_compare = 1
        Exit Function
    End If
    If a.pos.y <> b.pos.y Then
        If a.pos.y < b.pos.y Then Blueprint_compare = -1 Else Blueprint_compare = 1
        Exit Function
    End If
    If a.pos.z <> b.pos.z Then
        If a.pos.z < b.pos.z Then Blueprint_compare = -1 Else Blueprint_compare = 1
        Exit Function
    End If
    If a.dir_ <> b.dir_ Then
        If a.dir_ < b.dir_ Then Blueprint_compare = -1 Else Blueprint_compare = 1
        Exit Function
    End If
    If a.index <> b.index Then
        If ULT(a.index, b.index) Then Blueprint_compare = -1 Else Blueprint_compare = 1
        Exit Function
    End If
    If a.state <> b.state Then
        If a.state < b.state Then Blueprint_compare = -1 Else Blueprint_compare = 1
        Exit Function
    End If
    Blueprint_compare = 0
End Function

' optional<STRUCTURE> buildBlueprint() const: a structure handle, 0 = nullopt
Private Function Blueprint_buildBlueprint(b As Blueprint) As Long
    Blueprint_buildBlueprint = buildBlueprint(b.stats, b.pos, b.dir_, b.index, b.state, b.player)
End Function

Private Sub Blueprint_renderBlueprint(b As Blueprint, viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim optStruct As Long
    If clipXY(b.pos.x, b.pos.y) Then
        optStruct = Blueprint_buildBlueprint(b)
        If Not (optStruct <> 0) Then ASSERT_FAILURE "optStruct.has_value()", "src/display3d.cpp:368", "renderBlueprint", "buildBlueprint returned nullopt": Exit Sub
        renderStructure optStruct, viewMatrix, perspectiveViewMatrix
    End If
End Sub

'********************  Functions  ********************

Public Sub display3dScreenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    If psWScreen Is Nothing Then Exit Sub
    resizeRadar ' recalculate radar position
End Sub

Public Function radarVisible() As Boolean
    If radarOnScreen And radarPermitted And getWidgetsStatus() Then
        radarVisible = True
    Else
        radarVisible = False
    End If
End Function

Public Function interpolateAngleDegrees(ByVal a As Long, ByVal b As Long, ByVal t As Single) As Single
    Dim d As Single
    If a > 180 Then
        a = a - 360
    End If
    If b > 180 Then
        b = b - 360
    End If

    d = b - a

    interpolateAngleDegrees = a + d * t
End Function

' strImd: an iIMDShape * (WzGfx handle)
Public Function drawShape(ByVal strImd As Long, ByVal timeAnimationStarted As Long, ByVal colour As Long, ByVal buildingBrightness As Long, ByVal pieFlag As Long, ByVal pieFlagData As Long, modelMatrix As mat4, viewMatrix As mat4, Optional ByVal stretchDepth As Single = 0!) As Boolean
    Dim modifiedModelMatrix As mat4, animFrame_ As Long, elapsed As Long, frame As Long, objanimtime As Long, objanimframes As Long
    Dim state As ANIMFRAME, nextState As ANIMFRAME, frameFraction As Single, nextFrame As Long, q As Double, m As mat4
    modifiedModelMatrix = modelMatrix
    animFrame_ = 0 ' for texture animation
    If iIMDShape_numFrames(strImd) > 0 Then     ' Calculate an animation frame
        animFrame_ = getModularScaledGraphicsTime(iIMDShape_animInterval(strImd), iIMDShape_numFrames(strImd))
    End If
    objanimframes = iIMDShape_objanimframes(strImd)
    If objanimframes <> 0 Then
        objanimtime = iIMDShape_objanimtime(strImd)
        elapsed = graphicsTime - timeAnimationStarted
        If elapsed < 0 Then
            elapsed = 0 ' Animation hasn't started yet.
        End If

        frame = (elapsed \ objanimtime) Mod objanimframes
        If Not (frame < objanimframes) Then ASSERT_FAILURE "frame < strImd->objanimframes", "src/display3d.cpp:617", "drawShape", "Bad index %d >= %d", frame, objanimframes

        iIMDShape_objanimdata_p strImd, frame, VarPtr(state)

        If state.scale.x = -1! Then ' disabled frame, for implementing key frame animation
            drawShape = False
            Exit Function
        End If

        If iIMDShape_interpolate(strImd) = 1 Then
            ' static_cast<float>(fmod(elapsed / (float)strImd->objanimtime, strImd->objanimframes) - frame)
            q = CSng(elapsed / CSng(objanimtime))
            frameFraction = CSng((q - Fix(q / objanimframes) * objanimframes) - frame)
            nextFrame = (frame + 1) Mod objanimframes
            iIMDShape_objanimdata_p strImd, nextFrame, VarPtr(nextState)

            m = glm_interpolate(glm_translate(state.pos), glm_translate(nextState.pos), frameFraction)
            m = glm_mat4_mul(m, glm_rotate(RADIANS(interpolateAngleDegrees(state.rot.pitch \ DEG(1), nextState.rot.pitch \ DEG(1), frameFraction)), glm_vec3(1!, 0!, 0!)))
            m = glm_mat4_mul(m, glm_rotate(RADIANS(interpolateAngleDegrees(state.rot.direction \ DEG(1), nextState.rot.direction \ DEG(1), frameFraction)), glm_vec3(0!, 1!, 0!)))
            m = glm_mat4_mul(m, glm_rotate(RADIANS(interpolateAngleDegrees(state.rot.roll \ DEG(1), nextState.rot.roll \ DEG(1), frameFraction)), glm_vec3(0!, 0!, 1!)))
            m = glm_mat4_mul(m, glm_scale(state.scale))
            modifiedModelMatrix = glm_mat4_mul(modifiedModelMatrix, m)
        Else
            m = glm_translate(state.pos)
            m = glm_mat4_mul(m, glm_rotate(UNDEG(state.rot.pitch), glm_vec3(1!, 0!, 0!)))
            m = glm_mat4_mul(m, glm_rotate(UNDEG(state.rot.direction), glm_vec3(0!, 1!, 0!)))
            m = glm_mat4_mul(m, glm_rotate(UNDEG(state.rot.roll), glm_vec3(0!, 0!, 1!)))
            m = glm_mat4_mul(m, glm_scale(state.scale))
            modifiedModelMatrix = glm_mat4_mul(modifiedModelMatrix, m)
        End If
    End If

    drawShape = pie_Draw3DShape(strImd, animFrame_, colour, buildingBrightness, pieFlag, pieFlagData, VarPtr(modifiedModelMatrix), VarPtr(viewMatrix), stretchDepth, True)
End Function

Public Sub setScreenDispWithPerspective(sDisplay As SCREEN_DISP_DATA, perspectiveViewModelMatrix As mat4)
    Dim sx As Long, sy As Long
    ' Vector3i zero(0, 0, 0); Vector2i s(0, 0);
    pie_RotateProjectWithPerspective 0, 0, 0, VarPtr(perspectiveViewModelMatrix), sx, sy
    sDisplay.screenX = sx
    sDisplay.screenY = sy
End Sub

Public Sub setSkyBox(ByVal page As String, ByVal mywind As Single, ByVal myscale As Single)
    windSpeed = mywind
    wind = 0!
    skybox_scale = myscale
    pie_Skybox_Texture page
End Sub

' static inline void rotateSomething(int &x, int &y, uint16_t angle)
Private Sub rotateSomething(x As Long, y As Long, ByVal angle As Long)
    Dim cra As Double, sra As Double, newX As Long, newY As Long
    angle = angle And &HFFFF&
    cra = iCos(angle)
    sra = iSin(angle)
    ' int64_t products >> 16 (an arithmetic shift: the floor of the division)
    newX = Int((x * cra - y * sra) / 65536#)
    newY = Int((x * sra + y * cra) / 65536#)
    x = newX
    y = newY
End Sub

Private Function getTileBlueprint(ByVal mapX As Long, ByVal mapY As Long) As Blueprint
    Dim mouse As Vector2i, size As Vector2i, i As Long
    mouse.x = world_coord(mapX) + TILE_UNITS \ 2
    mouse.y = world_coord(mapY) + TILE_UNITS \ 2

    For i = 0 To blueprints_size - 1
        size = STRUCTURE_STATS_size(blueprints(i).stats, blueprints(i).dir_)
        size.x = size.x * TILE_UNITS
        size.y = size.y * TILE_UNITS
        If Abs(mouse.x - blueprints(i).pos.x) < size.x \ 2 And Abs(mouse.y - blueprints(i).pos.y) < size.y \ 2 Then
            getTileBlueprint = blueprints(i)
            Exit Function
        End If
    Next

    getTileBlueprint = Blueprint_(-1, Vector3i_(0, 0, 0), 0, 0, SS_BEING_BUILT, selectedPlayer)
End Function

' STRUCTURE *: a structure handle, valid until the next call (the C++'s static optional<STRUCTURE>)
Public Function getTileBlueprintStructure(ByVal mapX As Long, ByVal mapY As Long) As Long
    Static optStruct As Long            ' static optional<STRUCTURE> optStruct;
    Dim bp As Blueprint, b As Long

    bp = getTileBlueprint(mapX, mapY)
    If bp.state = SS_BLUEPRINT_PLANNED Then
        If optStruct <> 0 Then
            ' Delete previously returned structure, if any.
            optStruct = 0
        End If
        b = Blueprint_buildBlueprint(bp)
        ' Don't use `operator=(optional&&)` to avoid declaring custom copy/move-assignment operators
        ' for `STRUCTURE` (more specifically, for `SIMPLE_OBJECT`, which has some const-qualified data members).
        If b <> 0 Then
            ' (VB) the static's own copy: buildBlueprint's slot is overwritten by the next blueprint drawn
            optStruct = StructContainer_scratch(3)
            aStruct(optStruct) = aStruct(b)
        Else
            optStruct = 0
        End If
        getTileBlueprintStructure = optStruct  ' This blueprint was clicked on.
        Exit Function
    End If

    getTileBlueprintStructure = 0
End Function

' STRUCTURE_STATS const *: an index, -1 = nullptr
Public Function getTileBlueprintStats(ByVal mapX As Long, ByVal mapY As Long) As Long
    getTileBlueprintStats = getTileBlueprint(mapX, mapY).stats
End Function

Public Function anyBlueprintTooClose(ByVal stats As Long, pos As Vector2i, ByVal dir_ As Long) As Boolean
    Dim i As Long, bpos As Vector2i
    For i = 0 To blueprints_size - 1
        If blueprints(i).state = SS_BLUEPRINT_PLANNED Or blueprints(i).state = SS_BLUEPRINT_PLANNED_BY_ALLY Then
            bpos.x = blueprints(i).pos.x
            bpos.y = blueprints(i).pos.y
            If isBlueprintTooClose(stats, pos, dir_, blueprints(i).stats, bpos, blueprints(i).dir_) Then
                anyBlueprintTooClose = True
                Exit Function
            End If
        End If
    Next
    anyBlueprintTooClose = False
End Function

Public Sub clearBlueprints()
    blueprints_size = 0
    Erase blueprints
End Sub

' (VB) blueprints.push_back(b)
Private Sub blueprints_push_back(b As Blueprint)
    If blueprints_size = 0 Then
        ReDim blueprints(0 To 15)
    ElseIf blueprints_size > UBound(blueprints) Then
        ReDim Preserve blueprints(0 To 2 * blueprints_size - 1)
    End If
    blueprints(blueprints_size) = b
    blueprints_size = blueprints_size + 1
End Sub

' (VB) PIELIGHT's bytes
Private Function pielight_r(ByVal c As Long) As Long
    pielight_r = c And &HFF&
End Function

Private Function pielight_g(ByVal c As Long) As Long
    pielight_g = SHRU(c, 8) And &HFF&
End Function

Private Function pielight_b(ByVal c As Long) As Long
    pielight_b = SHRU(c, 16) And &HFF&
End Function

Private Function pielight_a(ByVal c As Long) As Long
    pielight_a = SHRU(c, 24) And &HFF&
End Function

Public Function selectionBrightness() As Long
    Dim brightVar As Long
    If Not gamePaused() Then
        brightVar = getModularScaledGraphicsTime(990, 110)
        If brightVar > 55 Then
            brightVar = 110 - brightVar
        End If
    Else
        brightVar = 55
    End If
    selectionBrightness = pal_SetBrightness(200 + brightVar)
End Function

Public Function structureBrightness(ByVal psStructure As Long) As Long
    Dim buildingBrightness As Long

    If STRUCTURE_isBlueprint(psStructure) Then
        buildingBrightness = getBlueprintColour(aStruct(psStructure).status)
    Else
        ' static_cast<UBYTE>(200 - 100 / 65536.f * getStructureDamage(psStructure))
        buildingBrightness = pal_SetBrightness(Fix(CSng(200! - CSng(100! / 65536!) * getStructureDamage(psStructure))) And &HFF&)

        ' If it's selected, then it's brighter
        If aStruct(psStructure).selected <> 0 Then
            buildingBrightness = selectionBrightness()
        End If
        If Not getRevealStatus() Then
            buildingBrightness = pal_SetBrightness(avGetObjLightLevel(psStructure, pielight_r(buildingBrightness)))
        End If
        If Not hasSensorOnTile(mapTile(map_coord(aStruct(psStructure).pos.x), map_coord(aStruct(psStructure).pos.y)), selectedPlayer) Then
            buildingBrightness = pal_RGBA(pielight_r(buildingBrightness) \ 2, pielight_g(buildingBrightness) \ 2, pielight_b(buildingBrightness) \ 2, pielight_a(buildingBrightness))
        End If
    End If
    structureBrightness = buildingBrightness
End Function


' Show all droid movement parts by displaying an explosion at every step
Private Sub showDroidPaths()
    Dim it As Long, psDroid As Long, length As Long, i As Long, pos As Vector3i
    If ((graphicsTime \ 250) Mod 2) <> 0 Then
        Exit Sub
    End If

    If selectedPlayer >= MAX_PLAYERS Then
        Exit Sub ' no-op for now
    End If

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected <> 0 And aDroid(psDroid).sMove.Status <> MOVEINACTIVE Then
            length = aDroid(psDroid).sMove.asPath_size
            i = aDroid(psDroid).sMove.pathIndex - 1
            If i < 0 Then i = 0
            Do While i < length
                If Not worldOnMap(aDroid(psDroid).sMove.asPath(i).x, aDroid(psDroid).sMove.asPath(i).y) Then ASSERT_FAILURE "worldOnMap(psDroid->sMove.asPath[i].x, psDroid->sMove.asPath[i].y)", "src/display3d.cpp:819", "showDroidPaths", "Path off map!"
                pos.x = aDroid(psDroid).sMove.asPath(i).x
                pos.z = aDroid(psDroid).sMove.asPath(i).y
                pos.y = map_Height(pos.x, pos.z) + 16

                effectGiveAuxVar 80
                addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_LASER, False, 0, 0
                i = i + 1
            Loop
        End If
        it = lnNext(it)
    Loop
End Sub

' Displays an image for the Network Issue button
' (VB) a W_BUTTON's display function: the widgets come with cluster G (setupConnectionStatusForm adds no button yet)
Private Sub NetworkDisplayImage(ByVal psWidget As Long, ByVal xOffset As Long, ByVal yOffset As Long)
End Sub

' (VB) in part: the connection status mask is computed as C++ does; the form and its buttons (widgAddForm,
' widgAddButton, widgDelete with NetworkDisplayImage) wait for the widgets, cluster G
Private Sub setupConnectionStatusForm()
    Static prevStatusMask As Long       ' unsigned = 0
    Dim statusMask As Long, total As Long, i As Long
    Const separation As Long = 3

    statusMask = 0
    total = 0
    For i = 0 To CONNECTIONSTATUS_NORMAL - 1
        If NETcheckPlayerConnectionStatus(i, NET_ALL_PLAYERS) Then
            statusMask = statusMask Or SHL(1, i)
            total = total + 1
        End If
    Next

    If prevStatusMask <> 0 And statusMask <> prevStatusMask Then
        ' Remove the icons. (VB: no widgets yet)
        prevStatusMask = 0
    End If

    If prevStatusMask = 0 And statusMask <> 0 Then
        ' Create the basic form and add the buttons (VB: no widgets yet)
        prevStatusMask = statusMask
    End If
End Sub

' Render the 3D world
Public Sub draw3DScene()
    Dim buildInfo As String, showMs As Boolean, modN As Long, modD As Long
    Dim trimMsg As String, width As Long, height As Long, visibleDroids As Long, undrawnDroids As Long, it As Long, psDroid As Long
    wzPerfBegin PERF_START_FRAME, "Start 3D scene"

    ' What frame number are we on?
    currentGameFrame = frameGetFrameNumber()

    ' Tell shader system what the time is
    pie_SetShaderTime graphicsTime

    ' Build the drag quad
    If dragBox3D.status = DRAG_RELEASED Then
        dragQuad.coords(0).x = dragBox3D.x1 ' TOP LEFT
        dragQuad.coords(0).y = dragBox3D.y1

        dragQuad.coords(1).x = dragBox3D.x2 ' TOP RIGHT
        dragQuad.coords(1).y = dragBox3D.y1

        dragQuad.coords(2).x = dragBox3D.x2 ' BOTTOM RIGHT
        dragQuad.coords(2).y = dragBox3D.y2

        dragQuad.coords(3).x = dragBox3D.x1 ' BOTTOM LEFT
        dragQuad.coords(3).y = dragBox3D.y2
    End If

    pie_Begin3DScene
    ' Set 3D world origins
    pie_SetGeometricOffset rendSurface_width \ 2, geoOffset

    updateFogDistance distance

    ' Set light manager
    If war_getPointLightPerPixelLighting() And getTerrainShaderQuality() = NORMAL_MAPPING Then
        setLightingManager renderingNew_LightingManager_new()
    Else
        setLightingManager New rendering1999_LightingManager
    End If

    ' Now, draw the terrain
    drawTiles playerPos

    wzPerfBegin PERF_MISC, "3D scene - misc and text"

    pie_BeginInterface

    ' Show the drag Box if necessary
    drawDragBox

    ' Have we released the drag box?
    If dragBox3D.status = DRAG_RELEASED Then
        dragBox3D.status = DRAG_INACTIVE
    End If

    drawDroidAndStructureSelections

    pie_SetFogStatus False
    iV_SetTextColour psPalette(WZCOL_TEXT_BRIGHT)

    ' Dont remove this folks!!!!
    If StrPtr(errorWaiting) <> 0 Then
        ' print the error message if none have been printed for one minute
        If lastErrorTime = 0 Or U32(lastErrorTime) + 60# * GAME_TICKS_PER_SEC < U32(realTime) Then
            audio_PlayBuildFailedOnce
            trimMsg = va("Error! (Check your logs!): %.78s", errorWaiting)
            If Len(trimMsg) > 254 Then trimMsg = Left$(trimMsg, 254)  ' char trimMsg[255]
            addConsoleMessage trimMsg, DEFAULT_JUSTIFY, NOTIFY_MESSAGE
            errorWaiting = vbNullString
            lastErrorTime = realTime
        End If
    Else
        errorWaiting = debugLastError()
    End If
    If showSAMPLES Then      ' Displays the number of sound samples we currently have
        txtShowSamples_Que.setText va("Que: %04u", audio_GetSampleQueueCount()), font_regular
        txtShowSamples_Lst.setText va("Lst: %04u", audio_GetSampleListCount()), font_regular
        txtShowSamples_Act.setText va("Act: %04u", sound_GetActiveSamplesCount()), font_regular

        width = txtShowSamples_Que.width() + 11
        height = txtShowSamples_Que.height()

        txtShowSamples_Que.render pie_GetVideoBufferWidth() - width, height + 2, psPalette(WZCOL_TEXT_BRIGHT)
        txtShowSamples_Lst.render pie_GetVideoBufferWidth() - width, height + 48, psPalette(WZCOL_TEXT_BRIGHT)
        txtShowSamples_Act.render pie_GetVideoBufferWidth() - width, height + 59, psPalette(WZCOL_TEXT_BRIGHT)
    End If
    If showFPS Then
        txtShowFPS.setText va("FPS: %d", frameRate()), font_regular
        width = txtShowFPS.width() + 10
        height = 9 'txtShowFPS.height();
        txtShowFPS.render pie_GetVideoBufferWidth() - width, pie_GetVideoBufferHeight() - height, psPalette(WZCOL_TEXT_BRIGHT)
    End If
    If showUNITCOUNT And selectedPlayer < MAX_PLAYERS Then
        txtUnits.setText va("Units: %u lost / %u built / %u killed", missionData.unitsLost, missionData.unitsBuilt, getSelectedPlayerUnitsKilled()), font_regular
        width = txtUnits.width() + 10
        height = 9 'txtUnits.height();
        If showFPS Then
            txtUnits.render pie_GetVideoBufferWidth() - width - (txtShowFPS.width() + 10), pie_GetVideoBufferHeight() - height, psPalette(WZCOL_TEXT_BRIGHT)
        Else
            txtUnits.render pie_GetVideoBufferWidth() - width, pie_GetVideoBufferHeight() - height, psPalette(WZCOL_TEXT_BRIGHT)
        End If
    End If
    If showORDERS Then
        txtShowOrders.setText DROIDDOING, font_regular
        height = txtShowOrders.height()
        txtShowOrders.render 0, pie_GetVideoBufferHeight() - height, psPalette(WZCOL_TEXT_BRIGHT)
    End If
    If showDROIDcounts And selectedPlayer < MAX_PLAYERS Then
        visibleDroids = 0
        undrawnDroids = 0
        it = apsDroidLists(selectedPlayer).head
        Do While it <> 0
            psDroid = lnObj(it)
            If aDroid(psDroid).sDisplay.frameNumber <> currentGameFrame Then
                undrawnDroids = undrawnDroids + 1
            Else
                visibleDroids = visibleDroids + 1
            End If
            it = lnNext(it)
        Loop
        droidText.setText va("Droids: %d drawn, %d undrawn", visibleDroids, undrawnDroids), font_regular
        droidText.render pie_GetVideoBufferWidth() - droidText.width() - 10, droidText.height() + 2, psPalette(WZCOL_TEXT_BRIGHT)
    End If

    setupConnectionStatusForm

    If getWidgetsStatus() And Not gamePaused() Then
        ' bool showMs = (gameTimeGetMod() < Rational(1, 4));
        gameTimeGetMod modN, modD
        showMs = (CDbl(modN) * 4# < CDbl(modD))
        If showMs Then
            buildInfo = getAsciiTime(graphicsTime, showMs)
        Else
            buildInfo = getAsciiTime(gameTime, showMs)
        End If
        txtLevelName.render RET_X + 134, 410 + (pie_GetVideoBufferHeight() - BASE_COORDS_Y), psPalette(WZCOL_TEXT_MEDIUM)
        If gInputManager_debugManager_debugMappingsAllowed() Then
            txtDebugStatus.render RET_X + 134, 436 + (pie_GetVideoBufferHeight() - BASE_COORDS_Y), psPalette(WZCOL_TEXT_MEDIUM)
        End If
        txtCurrentTime.setText buildInfo, font_small
        txtCurrentTime.render RET_X + 134, 422 + (pie_GetVideoBufferHeight() - BASE_COORDS_Y), psPalette(WZCOL_TEXT_MEDIUM)
    End If

    Do While playerPos.r.y > DEG(360)
        playerPos.r.y = playerPos.r.y - DEG(360)
    Loop

    ' If we don't have an active camera track, then track terrain height!
    If Not getWarCamStatus() Then
        ' Move the autonomous camera if necessary
        updatePlayerAverageCentreTerrainHeight
        trackHeight calculateCameraHeight(averageCentreTerrainHeight)
    Else
        processWarCam
    End If

    processSensorTarget
    processDestinationTarget

    structureEffects ' add fancy effects to structures

    showDroidSensorRanges 'shows sensor data for units/droids/whatever...-Q 5-10-05
    If CauseCrash Then
        ' (VB) writeGameInfo("WZdebuginfo.txt") and the crash handler's test crash (crashHandlingProviderTestCrash,
        ' the deliberate null-dereference) are not ported: the log and exit(-1)
        debug_ 1184, LOG_FATAL, "draw3DScene", "Forcing a segfault! (crash handler test)"
        exit_ -1 ' should never reach this, but just in case...
    End If
    'visualize radius if needed
    If bRangeDisplay Then
        drawRangeAtPos rangeCenterX, rangeCenterY, rangeRadius
    End If

    If showPath Then
        showDroidPaths
    End If

    wzPerfEnd PERF_MISC
End Sub


'**************************************************************************
Public Function doWeDrawProximitys() As Boolean
    doWeDrawProximitys = bDrawProximitys
End Function
'**************************************************************************
Public Sub setProximityDraw(ByVal val_ As Boolean)
    bDrawProximitys = val_
End Sub
'**************************************************************************
' Calculate the average terrain height for the area directly below the tile
Private Function calcAverageTerrainHeight(ByVal tileX As Long, ByVal tileZ As Long) As Long
    Dim numTilesAveraged As Long, result As Long, i As Long, j As Long, psTile As Long

    numTilesAveraged = 0

    '*
    ' * We track the height here - so make sure we get the average heights
    ' * of the tiles directly underneath us
    result = 0
    For i = -4 To 4
        For j = -4 To 4
            If tileOnMap(tileX + j, tileZ + i) Then
                ' Get a pointer to the tile at this location
                psTile = mapTile(tileX + j, tileZ + i)

                If psMapTiles(psTile).height > psMapTiles(psTile).waterLevel Then
                    result = result + psMapTiles(psTile).height
                Else
                    result = result + psMapTiles(psTile).waterLevel
                End If
                numTilesAveraged = numTilesAveraged + 1
            End If
        Next
    Next
    If numTilesAveraged = 0 Then ' might be if off map
        calcAverageTerrainHeight = ELEVATION_SCALE * TILE_UNITS
        Exit Function
    End If

    '*
    ' * Work out the average height.
    ' * We use this information to keep the player camera above the terrain.
    psTile = mapTile(tileX, tileZ)

    result = result \ numTilesAveraged
    If result < psMapTiles(psTile).height Then
        result = psMapTiles(psTile).height
    End If

    calcAverageTerrainHeight = result
End Function

Private Sub updatePlayerAverageCentreTerrainHeight()
    averageCentreTerrainHeight = calcAverageTerrainHeight(playerXTile, playerZTile)
End Sub

Private Function quadIntersectsWithScreen(quad_ As QUAD) As Boolean
    Dim width As Long, height As Long, iCoordIdx As Long
    width = pie_GetVideoBufferWidth()
    height = pie_GetVideoBufferHeight()
    For iCoordIdx = 0 To 3
        If (quad_.coords(iCoordIdx).x < 0) Or (quad_.coords(iCoordIdx).x > width) Then
            GoTo continue_
        End If
        If (quad_.coords(iCoordIdx).y < 0) Or (quad_.coords(iCoordIdx).y > height) Then
            GoTo continue_
        End If

        quadIntersectsWithScreen = True ' corner (x,y) is within the screen bounds
        Exit Function
continue_:
    Next

    quadIntersectsWithScreen = False
End Function

Public Function getBiasedShadowMapMVPMatrix(lightOrthoMatrix As mat4, lightViewMatrix As mat4) As mat4
    Dim biasMatrix As mat4, shadowMatrix As mat4
    ' glm::mat4 biasMatrix(0.5, 0.0, 0.0, 0.0,  0.0, 0.5, 0.0, 0.0,  0.0, 0.0, 0.5, 0.0,  0.5, 0.5, 0.5, 1.0) (column by column)
    biasMatrix.m(0) = 0.5!
    biasMatrix.m(5) = 0.5!
    biasMatrix.m(10) = 0.5!
    biasMatrix.m(12) = 0.5!
    biasMatrix.m(13) = 0.5!
    biasMatrix.m(14) = 0.5!
    biasMatrix.m(15) = 1!

    shadowMatrix = glm_mat4_mul(lightOrthoMatrix, lightViewMatrix)
    getBiasedShadowMapMVPMatrix = glm_mat4_mul(biasMatrix, shadowMatrix)
End Function

' Draw the terrain and all droids, missiles and other objects on it
' (VB) lightData, lightmap and lightManager are always getCurrentLightingData(), getCurrentLightmapData() and
' getCurrentLightingManager() (WzGfx's LightingData_*, LightMap_* and the manager setLightingManager set)
Private Sub drawTiles(player As iView)
    Dim currShadowMode As ShadowMode, baseViewMatrix As mat4, lightInvDir As Vector3f, numShadowCascades As Long
    Dim shadowCascades() As Cascade, shadowCascades_size As Long, viewMatrix As mat4, perspectiveMatrix As mat4
    Dim perspectiveViewMatrix As mat4, cameraPos As Vector3f, v4 As Vector4f, theSun As Vector3f, tileCalcPerspectiveViewMatrix As mat4
    Dim i As Long, j As Long, idx As Long, jdx As Long, sx As Long, sy As Long, pos As Vector3i, quad_ As QUAD
    Dim scaleF As Single, d As Long, shadowCascadesInfo_ As ShadowCascadesInfo, k As Long, lightManager As ILightingManager
    Dim m As mat4, w2s As constant_buffer_type_SHADER_WORLD_TO_SCREEN, negSun As Vector3f, lightmapMatrix As mat4
    Dim sunPos As Vector3f

    ' draw terrain

    currShadowMode = pie_getShadowMode()

    ' ----------------------------------------------------------------
    ' Do boundary and extent checking
    ' ----------------------------------------------------------------

    ' Find our position in tile coordinates
    playerXTile = map_coord(player.p.x)
    playerZTile = map_coord(player.p.z)

    ' ----------------------------------------------------------------
    ' Set up the geometry
    ' ----------------------------------------------------------------

    ' ----------------------------------------------------------------

    scaleF = pie_GetResScalingFactor() / 100!
    baseViewMatrix = glm_translate(glm_vec3(0!, 0!, distance))
    baseViewMatrix = glm_mat4_mul(baseViewMatrix, glm_scale(glm_vec3(scaleF, scaleF, scaleF)))
    baseViewMatrix = glm_mat4_mul(baseViewMatrix, glm_rotate(UNDEG(player.r.z), glm_vec3(0!, 0!, 1!)))
    baseViewMatrix = glm_mat4_mul(baseViewMatrix, glm_rotate(UNDEG(player.r.x), glm_vec3(1!, 0!, 0!)))
    baseViewMatrix = glm_mat4_mul(baseViewMatrix, glm_rotate(UNDEG(player.r.y), glm_vec3(0!, 1!, 0!)))
    baseViewMatrix = glm_mat4_mul(baseViewMatrix, glm_translate(glm_vec3(0!, CSng(-player.p.y), 0!)))

    ' Calculate shadow mapping cascades
    lightInvDir = getTheSun()
    numShadowCascades = gfx_api_context_numDepthPasses()
    If numShadowCascades > WZ_MAX_SHADOW_CASCADES Then numShadowCascades = WZ_MAX_SHADOW_CASCADES
    shadowCascades_size = 0
    If currShadowMode = ShadowMode_Shadow_Mapping Then
        calculateShadowCascades player, distance, baseViewMatrix, lightInvDir, numShadowCascades, shadowCascades, shadowCascades_size
    End If

    ' Incorporate all the view transforms into viewMatrix
    viewMatrix = glm_mat4_mul(baseViewMatrix, glm_translate(glm_vec3(CSng(-player.p.x), 0!, CSng(player.p.z))))

    pie_PerspectiveGet VarPtr(perspectiveMatrix)
    perspectiveViewMatrix = glm_mat4_mul(perspectiveMatrix, viewMatrix)

    v4 = glm_mat4_mul_vec4(glm_inverse(viewMatrix), glm_vec4(0!, 0!, 0!, 1!)) ' `actualCameraPosition` is not accurate enough due to int calc
    cameraPos = glm_vec3(v4.x, v4.y, v4.z)

    actualCameraPosition = Vector3i_(0, 0, 0)

    ' Set the camera position
    actualCameraPosition.z = actualCameraPosition.z - Fix(distance)

    ' Now, scale the world according to what resolution we're running in
    d = Fix(pie_GetResScalingFactor() / 100!)
    If d < 1 Then d = 1
    actualCameraPosition.z = actualCameraPosition.z \ d

    ' Rotate for the player
    rotateSomething actualCameraPosition.x, actualCameraPosition.y, -player.r.z
    rotateSomething actualCameraPosition.y, actualCameraPosition.z, -player.r.x
    rotateSomething actualCameraPosition.z, actualCameraPosition.x, -player.r.y

    ' Translate
    actualCameraPosition.y = actualCameraPosition.y - (-player.p.y)

    ' this also determines the length of the shadows
    v4 = glm_mat4_mul_vec4(viewMatrix, glm_vec4_v3(getTheSun(), 0!))
    theSun = glm_vec3(v4.x, v4.y, v4.z)
    pie_BeginLighting VarPtr(theSun)

    ' Reset all lighting data
    LightingData_lights_clear
    Set lightManager = getCurrentLightingManager()
    ILightingManager_SetFrameStart lightManager

    ' update the fog of war... FIXME: Remove this
    tileCalcPerspectiveViewMatrix = glm_mat4_mul(perspectiveMatrix, baseViewMatrix)
    idx = 0
    For i = -visibleTiles.y \ 2 To visibleTiles.y \ 2
        ' Go through the x's
        jdx = 0
        For j = -visibleTiles.x \ 2 To visibleTiles.x \ 2
            sx = 0
            sy = 0

            pos.x = world_coord(j) - (playerPos.p.x Mod TILE_WIDTH)
            pos.z = -world_coord(i) + (playerPos.p.z Mod TILE_HEIGHT)
            pos.y = 0

            If tileOnMap(playerXTile + j, playerZTile + i) Then
                pos.y = map_TileHeightSurface(playerXTile + j, playerZTile + i)
                LightMap_set playerXTile + j, playerZTile + i, pal_SetBrightness(0)
            End If
            tileScreenInfo(idx, jdx).z = pie_RotateProjectWithPerspective(pos.x, pos.y, pos.z, VarPtr(tileCalcPerspectiveViewMatrix), sx, sy)
            tileScreenInfo(idx, jdx).x = sx
            tileScreenInfo(idx, jdx).y = sy
            jdx = jdx + 1
        Next
        idx = idx + 1
    Next

    ' Determine whether each tile in the drawable range is actually visible on-screen
    ' (used for more accurate clipping elsewhere)
    For idx = 0 To visibleTiles.y - 1
        For jdx = 0 To visibleTiles.x - 1
            quad_.coords(0).x = tileScreenInfo(idx + 0, jdx + 0).x
            quad_.coords(0).y = tileScreenInfo(idx + 0, jdx + 0).y

            quad_.coords(1).x = tileScreenInfo(idx + 0, jdx + 1).x
            quad_.coords(1).y = tileScreenInfo(idx + 0, jdx + 1).y

            quad_.coords(2).x = tileScreenInfo(idx + 1, jdx + 1).x
            quad_.coords(2).y = tileScreenInfo(idx + 1, jdx + 1).y

            quad_.coords(3).x = tileScreenInfo(idx + 1, jdx + 0).x
            quad_.coords(3).y = tileScreenInfo(idx + 1, jdx + 0).y

            tileScreenVisible(idx, jdx) = quadIntersectsWithScreen(quad_)
        Next
    Next

    wzPerfEnd PERF_START_FRAME

    pie_StartMeshes

    ' This is done here as effects can light the terrain - pause mode problems though
    wzPerfBegin PERF_EFFECTS, "3D scene - effects"
    processEffects perspectiveViewMatrix
    atmosUpdateSystem
    avUpdateTiles
    wzPerfEnd PERF_EFFECTS

    ' The lightmap need to be ready at this point
    lightManager.ComputeFrameData VarPtr(perspectiveViewMatrix)

    ' prepare terrain for drawing
    perFrameTerrainUpdates

    ' and prepare for rendering the models
    wzPerfBegin PERF_MODEL_INIT, "Draw 3D scene - model init"

    ' ----------------------------------------------------------------
    ' Calculate & batch all mesh / object instances to be drawn
    ' ----------------------------------------------------------------
    displayStaticObjects viewMatrix, perspectiveViewMatrix ' may be bucket render implemented
    displayFeatures viewMatrix, perspectiveViewMatrix
    displayDynamicObjects viewMatrix, perspectiveViewMatrix ' may be bucket render implemented
    If doWeDrawProximitys() Then
        displayProximityMsgs viewMatrix, perspectiveViewMatrix
    End If
    displayDelivPoints viewMatrix, perspectiveViewMatrix
    display3DProjectiles viewMatrix, perspectiveViewMatrix ' may be bucket render implemented
    wzPerfEnd PERF_MODEL_INIT

    wzPerfBegin PERF_PARTICLES, "3D scene - particles"
    atmosDrawParticles viewMatrix, perspectiveViewMatrix
    wzPerfEnd PERF_PARTICLES

    bucketRenderCurrentList viewMatrix, perspectiveViewMatrix

    gfx_api_context_debugStringMarker "Draw 3D scene - blueprints"
    displayBlueprints viewMatrix, perspectiveViewMatrix

    ' ----------------------------------------------------------------
    ' Actually render / draw everything
    ' ----------------------------------------------------------------

    lightmapMatrix = getModelUVLightmapMatrix()
    pie_UpdateLightmap getTerrainLightmapTexture(), VarPtr(lightmapMatrix)
    pie_FinalizeMeshes currentGameFrame


    ' shadow/depth-mapping passes
    shadowCascadesInfo_ = ShadowCascadesInfo_new()       ' ShadowCascadesInfo shadowCascadesInfo;
    shadowCascadesInfo_.shadowMapSize = gfx_api_context_getDepthPassDimensions(0) ' Note: Currently assumes that every depth pass has the same dimensions
    For k = 0 To shadowCascades_size - 1
        If k >= WZ_MAX_SHADOW_CASCADES Then Exit For
        shadowCascadesInfo_.shadowMVPMatrix(k) = getBiasedShadowMapMVPMatrix(shadowCascades(k).projectionMatrix, shadowCascades(k).viewMatrix)
        shadowCascadesInfo_.shadowCascadeSplit(k) = shadowCascades(k).splitDepth
    Next

    If currShadowMode = ShadowMode_Shadow_Mapping Then
        For k = 0 To numShadowCascades - 1
            gfx_api_context_beginDepthPass k
            If bDrawTerrainShadows Then
                drawTerrainDepthOnly glm_mat4_mul(shadowCascades(k).projectionMatrix, shadowCascades(k).viewMatrix)
            End If
            pie_DrawAllMeshes currentGameFrame, VarPtr(shadowCascades(k).projectionMatrix), VarPtr(shadowCascades(k).viewMatrix), VarPtr(cameraPos), VarPtr(shadowCascadesInfo_), True
            gfx_api_context_endCurrentDepthPass
        Next
    End If
    ' start main render pass


    gfx_api_context_beginSceneRenderPass

    ' now we are about to draw the terrain
    wzPerfBegin PERF_TERRAIN, "3D scene - terrain"
    pie_SetFogStatus True
    sunPos = getTheSun()
    negSun = glm_vec3(-sunPos.x, -sunPos.y, -sunPos.z)
    drawTerrain perspectiveViewMatrix, viewMatrix, cameraPos, negSun, shadowCascadesInfo_
    wzPerfEnd PERF_TERRAIN

    ' draw skybox
    ' NOTE: Must come *after* drawTerrain *if* using the fallback (old) terrain shaders
    wzPerfBegin PERF_SKYBOX, "3D scene - skybox"
    pie_SkyboxPerspectiveGet VarPtr(m)
    renderSurroundings m, baseViewMatrix
    wzPerfEnd PERF_SKYBOX

    wzPerfBegin PERF_WATER, "3D scene - water"
    ' prepare for the water and the lightmap
    pie_SetFogStatus True
    ' also, make sure we can use world coordinates directly
    sunPos = getTheSun()
    negSun = glm_vec3(-sunPos.x, -sunPos.y, -sunPos.z)
    drawWater perspectiveViewMatrix, viewMatrix, cameraPos, negSun, shadowCascadesInfo_
    wzPerfEnd PERF_WATER

    wzPerfBegin PERF_MODELS, "3D scene - models"
    pie_DrawAllMeshes currentGameFrame, VarPtr(perspectiveMatrix), VarPtr(viewMatrix), VarPtr(cameraPos), VarPtr(shadowCascadesInfo_), False
    wzPerfEnd PERF_MODELS

    If Not gamePaused() Then
        doConstructionLines viewMatrix
    End If
    locateMouse

    gfx_api_context_endSceneRenderPass

    ' Draw the scene to the default framebuffer
    w2s.gamma = 1!
    gfx_api_pso_get(WorldToScreenPSO).bind
    gfx_api_pso_get(WorldToScreenPSO).bind_constants VarPtr(w2s), LenB(w2s)
    gfx_api_pso_get(WorldToScreenPSO).bind_vertex_buffers pScreenTriangleVBO
    gfx_api_pso_get(WorldToScreenPSO).bind_textures gfx_api_context_getSceneTexture()
    gfx_api_pso_get(WorldToScreenPSO).draw 3, 0
    gfx_api_pso_get(WorldToScreenPSO).unbind_vertex_buffers pScreenTriangleVBO
End Sub

' Initialise the fog, skybox and some other stuff
Public Function init3DView() As Boolean
    Dim sun As Vector3f, screenTriangleVertices(0 To 5) As Single

    getDefaultSunPosition sun.x, sun.y, sun.z
    setTheSun sun

    ' There are no drag boxes
    dragBox3D.status = DRAG_INACTIVE

    ' Get all the init stuff out of here?
    initWarCam

    ' Init the game messaging system
    initConsoleMessages

    atmosInitSystem

    ' default skybox, will override in script if not satisfactory
    setSkyBox "texpages/page-25-sky-arizona.png", 0!, 10000!

    ' distance is not saved, so initialise it now
    distance = war_GetMapZoom() ' distance

    If pie_GetFogEnabled() Then
        pie_SetFogStatus True
    End If

    ' Set the initial fog distance
    updateFogDistance distance

    setDefaultFogColour

    playerPos.r.z = 0 ' roll
    playerPos.r.y = 0 ' rotation
    playerPos.r.x = DEG(360 + INITIAL_STARTING_PITCH) ' angle

    If Not initTerrain() Then
        init3DView = False
        Exit Function
    End If

    txtLevelName.setText mapNameWithoutTechlevel(getLevelName()), font_small
    txtDebugStatus.setText "DEBUG ", font_small

    batchedObjectStatusRenderer_initialize


    If pScreenTriangleVBO Is Nothing Then
        ' vertex attributes for a single triangle that covers the screen
        screenTriangleVertices(0) = -1!: screenTriangleVertices(1) = -1!
        screenTriangleVertices(2) = 3!: screenTriangleVertices(3) = -1!
        screenTriangleVertices(4) = -1!: screenTriangleVertices(5) = 3!
        Set pScreenTriangleVBO = gfx_api_context_create_buffer_object(buffer_usage_vertex_buffer, buffer_storage_hint_static_draw, "screenTriangleVertices")
        pScreenTriangleVBO.upload 4 * 6, VarPtr(screenTriangleVertices(0))
    End If

    init3DView = True
End Function

Public Sub shutdown3DView()
    Set txtLevelName = New WzText
    Set txtDebugStatus = New WzText
    Set txtCurrentTime = New WzText
    Set txtShowFPS = New WzText
    Set txtUnits = New WzText
    ' show Samples text
    Set txtShowSamples_Que = New WzText
    Set txtShowSamples_Lst = New WzText
    Set txtShowSamples_Act = New WzText
    ' show Orders text
    Set txtShowOrders = New WzText
    ' show Droid visible/draw counts text
    Set droidText = New WzText

    batchedObjectStatusRenderer_clear ' NOTE: *NOT* reset() - see shutdown3DView_FullReset below for why
End Sub

Public Sub shutdown3DView_FullReset()
    ' Because of when levReleaseMissionData can be called (especially when using certain cheats in campaign like "ascend" or "let me win")
    ' Only call "reset" (which also resets BatchedMultiRectRenderer and actually frees GPU buffers)
    ' in this function, which is called from stageTwoShutdown (as opposed to stageThreeShutdown, which is called from levReleaseMissionData)...
    batchedObjectStatusRenderer_reset

    Set pScreenTriangleVBO = Nothing
End Sub

' set the view position from save game
Public Sub disp3d_setView(newView As iView)
    playerPos = newView
End Sub

' reset the camera rotation (used for save games <= 10)
Public Sub disp3d_oldView()
    playerPos.r.y = OLD_INITIAL_ROTATION ' rotation
    playerPos.p.y = OLD_START_HEIGHT ' height
End Sub

' get the view position for save game
Public Sub disp3d_getView(newView As iView)
    newView = playerPos
End Sub

' Are the current world coordinates within the processed range of tiles on the screen?
' (Warzone has a maximum range of tiles around the current player camera position that it will display.)
Public Function quickClipXYToMaximumTilesFromCurrentPosition(ByVal x As Long, ByVal y As Long) As Boolean
    ' +2 for edge of visibility fading (see terrain.cpp)
    If Abs(x - playerPos.p.x) < world_coord(visibleTiles.x \ 2 + 2) And _
       Abs(y - playerPos.p.z) < world_coord(visibleTiles.y \ 2 + 2) Then
        quickClipXYToMaximumTilesFromCurrentPosition = True
    Else
        quickClipXYToMaximumTilesFromCurrentPosition = False
    End If
End Function

' Are the current tile coordinates visible on screen?
Public Function clipXY(ByVal x As Long, ByVal y As Long) As Boolean
    Dim mapX As Long, mapY As Long
    ' +2 for edge of visibility fading (see terrain.cpp)
    If Abs(x - playerPos.p.x) < world_coord(visibleTiles.x \ 2 + 2) And _
       Abs(y - playerPos.p.z) < world_coord(visibleTiles.y \ 2 + 2) Then
        ' additional check using the tileScreenVisible matrix
        mapX = map_coord(x - playerPos.p.x) + visibleTiles.x \ 2
        mapY = map_coord(y - playerPos.p.z) + visibleTiles.y \ 2
        If mapX < 0 Or mapY < 0 Then
            clipXY = False
            Exit Function
        End If
        If mapX > visibleTiles.x Or mapY > visibleTiles.y Then
            clipXY = False
            Exit Function
        End If
        clipXY = tileScreenVisible(mapY, mapX)
    Else
        clipXY = False
    End If
End Function

Public Function clipXYZNormalized(normalizedPosition As Vector3i, perspectiveViewMatrix As mat4) As Boolean
    Dim px As Long, py As Long
    pie_RotateProjectWithPerspective normalizedPosition.x, normalizedPosition.y, normalizedPosition.z, VarPtr(perspectiveViewMatrix), px, py
    clipXYZNormalized = px >= 0 And py >= 0 And px < pie_GetVideoBufferWidth() And py < pie_GetVideoBufferHeight()
End Function

' Are the current 3d game-world coordinates visible on screen?
' (Does not take into account occlusion)
Public Function clipXYZ(ByVal x As Long, ByVal y As Long, ByVal z As Long, perspectiveViewMatrix As mat4) As Boolean
    Dim position As Vector3i
    position.x = x
    position.z = -(y)
    position.y = z

    clipXYZ = clipXYZNormalized(position, perspectiveViewMatrix)
End Function

' pIMD: an iIMDShape * (WzGfx handle, 0 = nullptr)
Public Function clipShapeOnScreen(ByVal pIMD As Long, perspectiveViewModelMatrix As mat4, Optional ByVal overdrawScreenPoints As Long = 10) As Boolean
    Dim wsRadius As Long, radius As Single, cZ As Single, cx As Long, cy As Long, screenMinX As Long, screenMinY As Long
    ' Get its absolute dimensions
    wsRadius = 22 ' World space radius, 22 = magic minimum

    If pIMD <> 0 Then
        If iIMDShape_radius(pIMD) > wsRadius Then wsRadius = iIMDShape_radius(pIMD)
    End If

    ' origin = Vector3i(0, wsRadius, 0); // take the center of the object

    ' get the screen coordinates
    cZ = pie_RotateProjectWithPerspective(0, wsRadius, 0, VarPtr(perspectiveViewModelMatrix), cx, cy) * 0.1!

    ' avoid division by zero
    If cZ > 0 Then
        radius = wsRadius / cZ * pie_GetResScalingFactor()
    Else
        radius = 1 ' 1 just in case some other code assumes radius != 0
    End If

    screenMinX = -overdrawScreenPoints
    screenMinY = -overdrawScreenPoints

    clipShapeOnScreen = (cx + radius > screenMinX And _
                         cx - radius < (pie_GetVideoBufferWidth() + overdrawScreenPoints) And _
                         cy + radius > screenMinY And _
                         cy - radius < (pie_GetVideoBufferHeight() + overdrawScreenPoints))
End Function

' Use overdrawScreenPoints as a workaround for casting shadows when the main unit is off-screen (but right at the edge)
Public Function clipDroidOnScreen(ByVal psDroid As Long, perspectiveViewModelMatrix As mat4, Optional ByVal overdrawScreenPoints As Long = 25) As Boolean
    Dim psBStats As Long, pIMD As Long, f As Long
    ' Get its absolute dimensions
    ' NOTE: This only takes into account body, but is "good enough"
    psBStats = DROID_getBodyStats(psDroid)
    pIMD = 0
    If psBStats >= 0 Then
        If asBodyStats(psBStats).pIMD <> 0 Then pIMD = iIMDBaseShape_displayModel(asBodyStats(psBStats).pIMD)
    End If

    If aDroid(psDroid).heightAboveMap > 0 Then
        ' clipShapeOnScreen is noticeably insufficient if it's a flying unit - we have to care about both shadows and the unit itself...
        ' HACK: boost the overdrawScreenPoints by a factor based on the heightAboveMap
        f = aDroid(psDroid).heightAboveMap \ 20
        If f < 1 Then f = 1
        overdrawScreenPoints = overdrawScreenPoints * f
    End If
    clipDroidOnScreen = clipShapeOnScreen(pIMD, perspectiveViewModelMatrix, overdrawScreenPoints)
End Function

Public Function clipStructureOnScreen(ByVal psStructure As Long) As Boolean
    Dim b As StructureBounds, start As Long, endB As Long, endW As Long, breadth As Long, width As Long
    b = getStructureBounds_struct(psStructure)
    ' assert(b.size.x != 0 && b.size.y != 0); (NDEBUG in a release build)
    ' +4 to make room for shadows on the terrain
    start = -2
    endB = (b.size.y + 2) - 1
    endW = (b.size.x + 2) - 1
    For breadth = start To endB
        For width = start To endW
            If width <> start And width <> endW And breadth <> start And breadth <> endB Then
                GoTo continue_ ' skip interior
            End If
            If clipXY(world_coord(b.map.x + width), world_coord(b.map.y + breadth)) Then
                clipStructureOnScreen = True
                Exit Function
            End If
continue_:
        Next
    Next

    clipStructureOnScreen = False
End Function

'* Get the screen coordinates for the current transform matrix.
' * This function is used to determine the area the user can click for the
' * intelligence screen buttons. The radius parameter is always set to the same value.
Public Sub calcFlagPosScreenCoords(pX As Long, pY As Long, pR As Long, perspectiveViewModelMatrix As mat4)
    Dim radius As Long, cx As Long, cy As Long
    ' Get it's absolute dimensions
    ' Vector3i center3d(0, 0, 0); Vector2i center2d(0, 0);
    ' How big a box do we want - will ultimately be calculated using xmax, ymax, zmax etc
    radius = 22

    ' Pop matrices and get the screen coordinates for last point
    pie_RotateProjectWithPerspective 0, 0, 0, VarPtr(perspectiveViewModelMatrix), cx, cy

    'store the coords
    pX = cx
    pY = cy
    pR = radius
End Sub

' Decide whether to render a projectile, and make sure it will be drawn
Private Sub display3DProjectiles(viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim psObj As Long, sub_ As WEAPON_SUBCLASS
    psObj = proj_GetFirst()
    Do While psObj <> 0
        ' If source or destination is visible, and projectile has been spawned and has not impacted.
        If Not ULT(graphicsTime, aProj(psObj).prevSpacetime.time_) And Not ULT(aProj(psObj).time_, graphicsTime) Then
            If gfxVisible(psObj) Then
                ' Draw a bullet at psObj->pos.x for X coord
                '   psObj->pos.y for Z coord
                '   whatever for Y (height) coord - arcing ?
                ' these guys get drawn last
                sub_ = asWeaponStats(aProj(psObj).psWStats).weaponSubClass
                If sub_ = WSC_ROCKET Or _
                   sub_ = WSC_MISSILE Or _
                   sub_ = WSC_COMMAND Or _
                   sub_ = WSC_SLOWMISSILE Or _
                   sub_ = WSC_SLOWROCKET Or _
                   sub_ = WSC_ENERGY Or _
                   sub_ = WSC_EMP Then
                    bucketAddTypeToList RENDER_PROJECTILE, psObj, perspectiveViewMatrix
                Else
                    renderProjectile psObj, viewMatrix, perspectiveViewMatrix
                End If
            End If
        End If

        psObj = proj_GetNext()
    Loop
End Sub ' end of function display3DProjectiles

' Draw a projectile to the screen
Public Sub renderProjectile(ByVal psCurr As Long, viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim psStats As Long, dv As Vector3i, pIMD As Long, st As Spacetime, sub_ As WEAPON_SUBCLASS
    Dim camera_base As Vector3i, camera_mod As Vector3i, modelMatrix_base As mat4, camera As Vector3i, modelMatrix As mat4
    Dim rollToCamera As Boolean, pitchToCamera As Boolean, premultiplied As Boolean, additive As Boolean, flags As Long
    Dim cx As Long, cy As Long, cz As Long, x As Long, z As Long

    psStats = aProj(psCurr).psWStats
    sub_ = asWeaponStats(psStats).weaponSubClass
    ' Reject flame or command since they have interim drawn fx
    If sub_ = WSC_FLAME Or _
       sub_ = WSC_COMMAND Or _
       sub_ = WSC_ELECTRONIC Or _
       sub_ = WSC_EMP Or _
       (bMultiPlayer And sub_ = WSC_LAS_SAT) Then
        ' We don't do projectiles from these guys, cos there's an effect instead
        Exit Sub
    End If

    st = interpolateObjectSpacetime(psCurr, graphicsTime)

    'the weapon stats holds the reference to which graphic to use
    'Need to draw the graphic depending on what the projectile is doing - hitting target,
    'missing target, in flight etc - JUST DO IN FLIGHT FOR NOW!
    pIMD = 0
    If asWeaponStats(psStats).pInFlightGraphic <> 0 Then pIMD = iIMDBaseShape_displayModel(asWeaponStats(psStats).pInFlightGraphic)

    If Not clipXYZ(st.pos.x, st.pos.y, st.pos.z, perspectiveViewMatrix) Then
        Exit Sub ' Projectile is not on the screen (Note: This uses the position point of the projectile, not a full shape clipping check, for speed.)
    End If

    ' Get bullet's x coord
    dv.x = st.pos.x

    ' Get it's y coord (z coord in the 3d world
    dv.z = -(st.pos.y)

    ' What's the present height of the bullet?
    dv.y = st.pos.z

    ' Set up the matrix
    camera_base = actualCameraPosition

    ' Translate to the correct position
    camera_mod.x = st.pos.x - playerPos.p.x
    camera_mod.z = -(st.pos.y - playerPos.p.z)
    camera_mod.y = st.pos.z
    camera_base.x = camera_base.x - camera_mod.x
    camera_base.y = camera_base.y - camera_mod.y
    camera_base.z = camera_base.z - camera_mod.z

    ' Rotate it to the direction it's facing
    rotateSomething camera_base.z, camera_base.x, -(-st.rot.direction)

    ' pitch it
    rotateSomething camera_base.y, camera_base.z, -st.rot.pitch

    modelMatrix_base = glm_translate(glm_vec3(CSng(dv.x), CSng(dv.y), CSng(dv.z)))
    modelMatrix_base = glm_mat4_mul(modelMatrix_base, glm_rotate(UNDEG(-st.rot.direction), glm_vec3(0!, 1!, 0!)))
    modelMatrix_base = glm_mat4_mul(modelMatrix_base, glm_rotate(UNDEG(st.rot.pitch), glm_vec3(1!, 0!, 0!)))

    Do While pIMD <> 0
        rollToCamera = False
        pitchToCamera = False
        premultiplied = False
        additive = (sub_ = WSC_ROCKET) Or (sub_ = WSC_MISSILE) Or (sub_ = WSC_SLOWROCKET) Or (sub_ = WSC_SLOWMISSILE)

        flags = iIMDShape_flags(pIMD)
        If (flags And iV_IMD_ROLL_TO_CAMERA) <> 0 Then
            rollToCamera = True
        End If
        If (flags And iV_IMD_PITCH_TO_CAMERA) <> 0 Then
            rollToCamera = True
            pitchToCamera = True
        End If
        If (flags And iV_IMD_NO_ADDITIVE) <> 0 Then
            additive = False
        End If
        If (flags And iV_IMD_ADDITIVE) <> 0 Then
            additive = True
        End If
        If (flags And iV_IMD_PREMULTIPLIED) <> 0 Then
            additive = False
            premultiplied = True
        End If

        camera = camera_base
        modelMatrix = modelMatrix_base

        If pitchToCamera Or rollToCamera Then
            ' Centre on projectile (relevant for twin projectiles).
            iIMDShape_connector pIMD, 0, cx, cy, cz
            camera.x = camera.x - cx
            camera.y = camera.y - cy
            camera.z = camera.z - cz
            modelMatrix = glm_mat4_mul(modelMatrix, glm_translate(glm_vec3(CSng(cx), CSng(cy), CSng(cz))))
        End If

        If pitchToCamera Then
            x = iAtan2(camera.z, camera.y)
            rotateSomething camera.y, camera.z, -x
            modelMatrix = glm_mat4_mul(modelMatrix, glm_rotate(UNDEG(x), glm_vec3(1!, 0!, 0!)))
        End If

        If rollToCamera Then
            z = -iAtan2(camera.x, camera.y)
            rotateSomething camera.x, camera.y, -z
            modelMatrix = glm_mat4_mul(modelMatrix, glm_rotate(UNDEG(z), glm_vec3(0!, 0!, 1!)))
        End If

        If pitchToCamera Or rollToCamera Then
            camera.x = camera.x - (-cx)
            camera.y = camera.y - (-cy)
            camera.z = camera.z - (-cz)
            ' Undo centre on projectile (relevant for twin projectiles).
            modelMatrix = glm_mat4_mul(modelMatrix, glm_translate(glm_vec3(CSng(-cx), CSng(-cy), CSng(-cz))))
        End If

        If premultiplied Then
            pie_Draw3DShape pIMD, 0, 0, psPalette(WZCOL_WHITE), pie_PREMULTIPLIED, 0, VarPtr(modelMatrix), VarPtr(viewMatrix), 0!, True
        ElseIf additive Then
            pie_Draw3DShape pIMD, 0, 0, psPalette(WZCOL_WHITE), pie_ADDITIVE, 164, VarPtr(modelMatrix), VarPtr(viewMatrix), 0!, True
        Else
            pie_Draw3DShape pIMD, 0, 0, psPalette(WZCOL_WHITE), 0, 0, VarPtr(modelMatrix), VarPtr(viewMatrix), 0!, True
        End If
        pIMD = iIMDShape_next(pIMD)
    Loop
End Sub

' Draw the buildings
Private Sub displayStaticObjects(viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim aPlayer As Long, it As Long, obj As Long, psStructure As Long, died As Long
    ' to solve the flickering edges of baseplates
'   pie_SetDepthOffset(-1.0f);

    ' Go through all the players
    For aPlayer = 0 To MAX_PLAYERS - 1
        ' Now go all buildings for that player
        it = apsStructLists(aPlayer).head
        Do While it <> 0
            obj = lnObj(it)
            ' Worth rendering the structure?
            died = SIMPLE_OBJECT_died(obj)
            If SIMPLE_OBJECT_type(obj) <> OBJ_STRUCTURE Then GoTo next_struct
            If died <> 0 And ULT(died, graphicsTime) Then GoTo next_struct
            If Not quickClipXYToMaximumTilesFromCurrentPosition(SIMPLE_OBJECT_pos(obj).x, SIMPLE_OBJECT_pos(obj).y) Then GoTo next_struct
            psStructure = castStructure(obj)

            If Not clipStructureOnScreen(psStructure) Then
                GoTo next_struct
            End If

            renderStructure psStructure, viewMatrix, perspectiveViewMatrix
next_struct:
            it = lnNext(it)
        Loop
    Next

    ' Walk through destroyed objects.

    ' Now go all buildings for that player
    it = psDestroyedObj.head
    Do While it <> 0
        obj = lnObj(it)
        ' Worth rendering the structure?
        died = SIMPLE_OBJECT_died(obj)
        If SIMPLE_OBJECT_type(obj) <> OBJ_STRUCTURE Then GoTo next_destroyed
        If died <> 0 And ULT(died, graphicsTime) Then GoTo next_destroyed
        If Not quickClipXYToMaximumTilesFromCurrentPosition(SIMPLE_OBJECT_pos(obj).x, SIMPLE_OBJECT_pos(obj).y) Then GoTo next_destroyed
        psStructure = castStructure(obj)

        If Not clipStructureOnScreen(psStructure) Then
            GoTo next_destroyed
        End If

        renderStructure psStructure, viewMatrix, perspectiveViewMatrix
next_destroyed:
        it = lnNext(it)
    Loop

'   pie_SetDepthOffset(0.0f);
End Sub

' tile: a MAPTILE * (psMapTiles' index); stats: a STRUCTURE_STATS *
Private Function tileHasIncompatibleStructure(ByVal tile As Long, ByVal stats As Long, ByVal moduleIndex As Long) As Boolean
    Dim psStruct As Long
    psStruct = castStructure(psMapTiles(tile).psObject)
    If psStruct = 0 Then
        tileHasIncompatibleStructure = False
        Exit Function
    End If
    If aStruct(psStruct).status = SS_BEING_BUILT Then
        If nextModuleToBuild(psStruct, -1) >= moduleIndex Then
            tileHasIncompatibleStructure = True
            Exit Function
        End If
    End If
    If isWall(asStructureStats(aStruct(psStruct).pStructureType).typ) And isBuildableOnWalls(asStructureStats(stats).typ) Then
        tileHasIncompatibleStructure = False
        Exit Function
    End If
    If IsStatExpansionModule(stats) Then
        tileHasIncompatibleStructure = False
        Exit Function
    End If
    tileHasIncompatibleStructure = True
End Function

Private Sub drawLineBuild(ByVal player As Long, ByVal psStats As Long, pos As Vector2i, pos2 As Vector2i, ByVal direction As Long, ByVal state As STRUCT_STATES)
    Dim lb As LineBuild, i As Long, cur As Vector2i, b As StructureBounds, z As Long, j As Long, k As Long, bp As Blueprint
    lb = calcLineBuild(psStats, direction, pos, pos2)

    For i = 0 To lb.count - 1
        cur = LineBuild_at(lb, i)
        If tileHasIncompatibleStructure(worldTile(cur.x, cur.y), psStats, 0) Then
            GoTo continue_  ' construction has started
        End If

        b = getStructureBounds_structStats(psStats, cur, direction)
        z = 0
        For j = 0 To b.size.y
            For k = 0 To b.size.x
                If map_TileHeight(b.map.x + k, b.map.y + j) > z Then z = map_TileHeight(b.map.x + k, b.map.y + j)
            Next
        Next
        bp = Blueprint_(psStats, Vector3i_(cur.x, cur.y, z), snapDirection(direction), 0, state, player) ' snapDirection may be unnecessary here
        blueprints_push_back bp
continue_:
    Next
End Sub

Private Sub renderBuildOrder(ByVal droidPlayer As Long, order As DROID_ORDER_DATA, ByVal state As STRUCT_STATES)
    Dim stats As Long, pos As Vector2i, structure_ As Long, b As StructureBounds, z As Long, i As Long, j As Long, bp As Blueprint
    pos = order.pos
    If order.typ = DORDER_BUILDMODULE Then
        structure_ = castStructure(order.psObj)
        If structure_ = 0 Then
            Exit Sub
        End If
        stats = getModuleStat(structure_)
        pos.x = aStruct(structure_).pos.x
        pos.y = aStruct(structure_).pos.y
    Else
        stats = order.psStats
    End If

    If stats < 0 Then
        Exit Sub
    End If

    'draw the current build site if its a line of structures
    If order.typ = DORDER_LINEBUILD Then
        drawLineBuild droidPlayer, stats, pos, order.pos2, order.direction, state
    End If
    If order.typ = DORDER_BUILD Or order.typ = DORDER_BUILDMODULE Then
        If Not tileHasIncompatibleStructure(mapTile(map_coord(pos.x), map_coord(pos.y)), stats, order.index) Then
            b = getStructureBounds_structStats(stats, pos, order.direction)
            z = 0
            For j = 0 To b.size.y
                For i = 0 To b.size.x
                    If map_TileHeight(b.map.x + i, b.map.y + j) > z Then z = map_TileHeight(b.map.x + i, b.map.y + j)
                Next
            Next
            bp = Blueprint_(stats, Vector3i_(pos.x, pos.y, z), snapDirection(order.direction), order.index, state, droidPlayer)
            blueprints_push_back bp
        End If
    End If
End Sub

Public Sub displayBlueprints(viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim state As STRUCT_STATES, stats As Long, direction As Long, width As Long, height As Long, pos As Vector2i
    Dim b As StructureBounds, z As Long, i As Long, j As Long, player As Long, it As Long, psDroid As Long, order As Long
    Dim tmp As Blueprint, n As Long

    clearBlueprints ' Delete old blueprints and draw new ones.

    If (buildState = BUILD3D_VALID Or buildState = BUILD3D_POS) And _
       sBuildDetails.x > 0 And sBuildDetails.x < mapWidth And _
       sBuildDetails.y > 0 And sBuildDetails.y < mapHeight Then
        If buildState = BUILD3D_VALID Then
            state = SS_BLUEPRINT_VALID
        Else
            state = SS_BLUEPRINT_INVALID
        End If
        ' we are placing a building or a delivery point
        stats = castStructureStats(sBuildDetails.psStats)
        If stats >= 0 Then
            ' it's a building
            direction = getBuildingDirection()
            If wallDrag.status = DRAG_PLACING Or wallDrag.status = DRAG_DRAGGING Then
                drawLineBuild selectedPlayer, stats, wallDrag.pos, wallDrag.pos2, direction, state
            Else
                If (direction And &H4000&) = 0 Then
                    width = sBuildDetails.width
                    height = sBuildDetails.height
                Else
                    ' Rotated 90 degrees, swap width and height
                    width = sBuildDetails.height
                    height = sBuildDetails.width
                End If
                ' a single building
                pos.x = world_coord(sBuildDetails.x) + world_coord(width) \ 2
                pos.y = world_coord(sBuildDetails.y) + world_coord(height) \ 2

                b = getStructureBounds_structStats(stats, pos, direction)
                z = 0
                For j = 0 To b.size.y
                    For i = 0 To b.size.x
                        If map_TileHeight(b.map.x + i, b.map.y + j) > z Then z = map_TileHeight(b.map.x + i, b.map.y + j)
                    Next
                Next

                If Not ValueTracker_isTracking(playerBlueprintX) Then
                    ValueTracker_startTracking playerBlueprintX, pos.x
                    ValueTracker_setSpeed playerBlueprintX, BlueprintTrackAnimationSpeed
                    ValueTracker_startTracking playerBlueprintY, pos.y
                    ValueTracker_setSpeed playerBlueprintY, BlueprintTrackAnimationSpeed
                    ValueTracker_startTracking playerBlueprintZ, z
                    ValueTracker_setSpeed playerBlueprintZ, BlueprintTrackAnimationSpeed
                    ValueTracker_startTracking playerBlueprintDirection, direction
                    ValueTracker_setSpeed playerBlueprintDirection, BlueprintTrackAnimationSpeed + 30
                End If

                ValueTracker_setTarget playerBlueprintX, pos.x
                ValueTracker_update playerBlueprintX
                ValueTracker_setTarget playerBlueprintY, pos.y
                ValueTracker_update playerBlueprintY
                ValueTracker_setTarget playerBlueprintZ, z
                ValueTracker_update playerBlueprintZ

                If ValueTracker_reachedTarget(playerBlueprintDirection) Then
                    ValueTracker_startTracking playerBlueprintDirection, ValueTracker_getTarget(playerBlueprintDirection)
                    ValueTracker_setTargetDelta playerBlueprintDirection, ToInt16(direction - ValueTracker_getTarget(playerBlueprintDirection))
                End If

                ValueTracker_update playerBlueprintDirection

                playerBlueprint.stats = stats
                playerBlueprint.pos.x = ValueTracker_getCurrent(playerBlueprintX)
                playerBlueprint.pos.y = ValueTracker_getCurrent(playerBlueprintY)
                playerBlueprint.pos.z = ValueTracker_getCurrent(playerBlueprintZ)
                playerBlueprint.dir_ = ValueTracker_getCurrent(playerBlueprintDirection) And &HFFFF&
                playerBlueprint.index = 0
                playerBlueprint.state = state
                playerBlueprint.player = selectedPlayer And &HFF&

                blueprints_push_back playerBlueprint
            End If
        End If
    Else
        ValueTracker_stopTracking playerBlueprintX
        ValueTracker_stopTracking playerBlueprintY
        ValueTracker_stopTracking playerBlueprintZ
        ValueTracker_stopTracking playerBlueprintDirection
    End If

    ' now we draw the blueprints for all ordered buildings
    For player = 0 To MAX_PLAYERS - 1
        If Not hasSharedVision(selectedPlayer, player) And Not NetPlay_players_isSpectator(selectedPlayer) Then
            GoTo next_player
        End If
        If player = selectedPlayer Then state = SS_BLUEPRINT_PLANNED Else state = SS_BLUEPRINT_PLANNED_BY_ALLY
        it = apsDroidLists(player).head
        Do While it <> 0
            psDroid = lnObj(it)
            If aDroid(psDroid).droidType = DROID_CONSTRUCT Or aDroid(psDroid).droidType = DROID_CYBORG_CONSTRUCT Then
                renderBuildOrder aDroid(psDroid).player, aDroid(psDroid).order, state
                'now look thru' the list of orders to see if more building sites
                For order = aDroid(psDroid).listPendingBegin To aDroid(psDroid).asOrderList_size - 1
                    renderBuildOrder aDroid(psDroid).player, aDroid(psDroid).asOrderList(order), state
                Next
            End If
            it = lnNext(it)
        Loop
next_player:
    Next

    ' Erase duplicate blueprints.
    ' std::sort(blueprints.begin(), blueprints.end()) (VB: an insertion sort; equal blueprints are erased next)
    For i = 1 To blueprints_size - 1
        tmp = blueprints(i)
        j = i - 1
        Do While j >= 0
            If Blueprint_compare(blueprints(j), tmp) <= 0 Then Exit Do
            blueprints(j + 1) = blueprints(j)
            j = j - 1
        Loop
        blueprints(j + 1) = tmp
    Next
    ' blueprints.erase(std::unique(blueprints.begin(), blueprints.end()), blueprints.end());
    If blueprints_size > 0 Then
        n = 1
        For i = 1 To blueprints_size - 1
            If Blueprint_compare(blueprints(n - 1), blueprints(i)) <> 0 Then
                blueprints(n) = blueprints(i)
                n = n + 1
            End If
        Next
        blueprints_size = n
    End If

    ' Actually render everything.
    For i = 0 To blueprints_size - 1
        Blueprint_renderBlueprint blueprints(i), viewMatrix, perspectiveViewMatrix
    Next

    renderDeliveryRepos viewMatrix, perspectiveViewMatrix
End Sub

' Draw Factory Delivery Points
Private Sub displayDelivPoints(viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim it As Long, psDelivPoint As Long
    If selectedPlayer >= MAX_PLAYERS Then Exit Sub ' no-op
    it = apsFlagPosLists(selectedPlayer).head
    Do While it <> 0
        psDelivPoint = lnObj(it)
        If clipXY(aFlagPos(psDelivPoint).coords.x, aFlagPos(psDelivPoint).coords.y) Then
            renderDeliveryPoint psDelivPoint, False, viewMatrix, perspectiveViewMatrix
        End If
        it = lnNext(it)
    Loop
End Sub

Public Function clipFeatureOnScreen(ByVal psFeature As Long) As Boolean
    Dim b As StructureBounds, overdrawTiles As Long, start As Long, endB As Long, endW As Long, breadth As Long, width As Long
    b = getStructureBounds(psFeature)
    ' assert(b.size.x != 0 && b.size.y != 0); (NDEBUG in a release build)
    If b.size.x < 2 And b.size.y < 2 Then overdrawTiles = 1 Else overdrawTiles = 2 ' make room for shadows on the terrain
    start = -overdrawTiles
    endB = (b.size.y + overdrawTiles) - 1
    endW = (b.size.x + overdrawTiles) - 1
    For breadth = start To endB
        For width = start To endW
            If width <> start And width <> endW And breadth <> start And breadth <> endB Then
                GoTo continue_ ' skip interior
            End If
            If clipXY(world_coord(b.map.x + width), world_coord(b.map.y + breadth)) Then
                clipFeatureOnScreen = True
                Exit Function
            End If
continue_:
        Next
    Next

    clipFeatureOnScreen = False
End Function

' Draw the features
Private Sub displayFeatures(viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim it As Long, obj As Long, psFeature As Long, died As Long
    ' player can only be 0 for the features.

    ' Go through all the features
    it = apsFeatureList(0).head
    Do While it <> 0
        obj = lnObj(it)
        If SIMPLE_OBJECT_type(obj) = OBJ_FEATURE Then
            died = SIMPLE_OBJECT_died(obj)
            If died = 0 Or ULT(graphicsTime, died) Then
                If quickClipXYToMaximumTilesFromCurrentPosition(SIMPLE_OBJECT_pos(obj).x, SIMPLE_OBJECT_pos(obj).y) Then
                    psFeature = castFeature(obj)
                    If clipFeatureOnScreen(psFeature) Then
                        renderFeature psFeature, viewMatrix, perspectiveViewMatrix
                    End If
                End If
            End If
        End If
        it = lnNext(it)
    Loop

    ' Walk through destroyed objects.

    ' Go through all the features
    it = psDestroyedObj.head
    Do While it <> 0
        obj = lnObj(it)
        If SIMPLE_OBJECT_type(obj) = OBJ_FEATURE Then
            died = SIMPLE_OBJECT_died(obj)
            If died = 0 Or ULT(graphicsTime, died) Then
                If clipXY(SIMPLE_OBJECT_pos(obj).x, SIMPLE_OBJECT_pos(obj).y) Then
                    psFeature = castFeature(obj)
                    renderFeature psFeature, viewMatrix, perspectiveViewMatrix
                End If
            End If
        End If
        it = lnNext(it)
    Loop
End Sub

' Draw the Proximity messages for the *SELECTED PLAYER ONLY*
Private Sub displayProximityMsgs(viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim it As Long, psProxDisp As Long, psMessage As Long, vd As Long, x As Long, y As Long
    If selectedPlayer >= MAX_PLAYERS Then Exit Sub ' no-op

    ' Go through all the proximity Displays
    it = apsProxDisp(selectedPlayer).head
    Do While it <> 0
        psProxDisp = lnObj(it)
        psMessage = aProxDisp(psProxDisp).psMessage
        If Not aMessage(psMessage).read Then
            If aProxDisp(psProxDisp).typ = POS_PROXDATA Then
                vd = aMessage(psMessage).pViewData
                If vd = 0 Then
                    GoTo continue_
                End If
                ' VIEW_PROXIMITY *pViewProximity = (VIEW_PROXIMITY *)psProxDisp->psMessage->pViewData->pData;
                If Not aViewData(vd).hasProxData Then
                    GoTo continue_
                End If
                x = aViewData(vd).prox_x
                y = aViewData(vd).prox_y
            Else
                If aMessage(psMessage).psObj = 0 Then
                    GoTo continue_    ' sanity check
                End If
                x = SIMPLE_OBJECT_pos(aMessage(psMessage).psObj).x
                y = SIMPLE_OBJECT_pos(aMessage(psMessage).psObj).y
            End If
            ' Is the Message worth rendering?
            If clipXY(x, y) Then
                renderProximityMsg psProxDisp, viewMatrix, perspectiveViewMatrix
            End If
        End If
continue_:
        it = lnNext(it)
    Loop
End Sub

' Draw the droids
Private Sub displayDynamicObjects(viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim player As Long, it As Long, psDroid As Long, obj As Long, died As Long
    ' Need to go through all the droid lists
    For player = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(player).head
        Do While it <> 0
            psDroid = lnObj(it)
            If psDroid = 0 Then GoTo next_droid
            died = aDroid(psDroid).died
            If died <> 0 And ULT(died, graphicsTime) Then GoTo next_droid
            If Not quickClipXYToMaximumTilesFromCurrentPosition(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y) Then GoTo next_droid

            ' No point in adding it if you can't see it?
            If BASE_OBJECT_visibleForLocalDisplay(psDroid) <> 0 Then
                displayComponentObject psDroid, viewMatrix, perspectiveViewMatrix
            End If
next_droid:
            it = lnNext(it)
        Loop
    Next

    ' Walk through destroyed objects.
    it = psDestroyedObj.head
    Do While it <> 0
        obj = lnObj(it)
        psDroid = castDroid(obj)
        If psDroid = 0 Then GoTo next_destroyed
        died = SIMPLE_OBJECT_died(obj)
        If died <> 0 And ULT(died, graphicsTime) Then GoTo next_destroyed
        If Not quickClipXYToMaximumTilesFromCurrentPosition(SIMPLE_OBJECT_pos(obj).x, SIMPLE_OBJECT_pos(obj).y) Then GoTo next_destroyed

        ' No point in adding it if you can't see it?
        If BASE_OBJECT_visibleForLocalDisplay(psDroid) <> 0 Then
            displayComponentObject psDroid, viewMatrix, perspectiveViewMatrix
        End If
next_destroyed:
        it = lnNext(it)
    Loop
End Sub

' Sets the player's position and view angle - defaults player rotations as well
Public Sub setViewPos(ByVal x As Long, ByVal y As Long, ByVal Pan As Boolean)
    playerPos.p.x = world_coord(x)
    playerPos.p.z = world_coord(y)
    playerPos.r.z = 0

    updatePlayerAverageCentreTerrainHeight

    If playerPos.p.y < averageCentreTerrainHeight Then
        playerPos.p.y = averageCentreTerrainHeight + CAMERA_PIVOT_HEIGHT - HEIGHT_TRACK_INCREMENTS
    End If

    If getWarCamStatus() Then
        camToggleStatus
    End If
End Sub

' Get the player position
Public Function getPlayerPos() As Vector2i
    getPlayerPos.x = playerPos.p.x
    getPlayerPos.y = playerPos.p.z
End Function

' Set the player position
Public Sub setPlayerPos(ByVal x As Long, ByVal y As Long)
    If Not (x >= 0 And x < world_coord(mapWidth) And y >= 0 And y < world_coord(mapHeight)) Then ASSERT_FAILURE "x >= 0 && x < world_coord(mapWidth) && y >= 0 && y < world_coord(mapHeight)", "src/display3d.cpp:2473", "setPlayerPos", "Position off map"
    playerPos.p.x = x
    playerPos.p.z = y
    playerPos.r.z = 0
End Sub

' Get the distance at which the player views the world
Public Function getViewDistance() As Single
    getViewDistance = distance
End Function

Public Function getIsCloseDistance() As Boolean
    getIsCloseDistance = bIsCloseDistance
End Function

' Set the distance at which the player views the world
Public Sub setViewDistance(ByVal dist As Single)
    distance = dist
    bIsCloseDistance = (distance < 750!)
    debug_ 2495, LOG_WZ, "setViewDistance", "Setting zoom to %.0f", distance
End Sub

Public Sub calcScreenCoords(ByVal psDroid As Long, perspectiveViewMatrix As mat4)
    Dim psBStats As Long, wsRadius As Long, radius As Single, cZ As Single, center As Vector2i
    ' Get it's absolute dimensions
    psBStats = DROID_getBodyStats(psDroid)
    wsRadius = 22 ' World space radius, 22 = magic minimum

    ' NOTE: This only takes into account body, but seems "good enough"
    If psBStats >= 0 Then
        If asBodyStats(psBStats).pIMD <> 0 Then
            If iIMDBaseShape_radius(asBodyStats(psBStats).pIMD) > wsRadius Then wsRadius = iIMDBaseShape_radius(asBodyStats(psBStats).pIMD)
        End If
    End If

    ' origin = Vector3i(0, wsRadius, 0); // take the center of the object

    ' get the screen coordinates
    cZ = pie_RotateProjectWithPerspective(0, wsRadius, 0, VarPtr(perspectiveViewMatrix), center.x, center.y) * 0.1!

    ' avoid division by zero
    If cZ > 0 Then
        radius = wsRadius / cZ * pie_GetResScalingFactor()
    Else
        radius = 1 ' 1 just in case some other code assumes radius != 0
    End If

    ' Deselect all the droids if we've released the drag box
    If dragBox3D.status = DRAG_RELEASED Then
        If inQuad(center, dragQuad) And aDroid(psDroid).player = selectedPlayer Then
            'don't allow Transporter Droids to be selected here
            'unless we're in multiPlayer mode!!!!
            If Not DROID_isTransporter(psDroid) Or bMultiPlayer Then
                dealWithDroidSelect psDroid, True
            End If
        End If
    End If

    ' Store away the screen coordinates so we can select the droids without doing a transform
    aDroid(psDroid).sDisplay.screenX = center.x
    aDroid(psDroid).sDisplay.screenY = center.y
    aDroid(psDroid).sDisplay.screenR = Fix(radius)
End Sub

Public Sub screenCoordToWorld(screenCoord As Vector2i, worldCoord As Vector2i, tileX As Long, tileY As Long)
    Dim nearestZ As Long, outMousePos As Vector2i, i As Long, j As Long, idx As Long, jdx As Long, tileZ As Long
    Dim quad_ As QUAD, piq As Vector2i
    nearestZ = INT32_MAX
    ' Intentionally not the same range as in drawTiles()
    idx = 0
    For i = -visibleTiles.y \ 2 To visibleTiles.y \ 2 - 1
        jdx = 0
        For j = -visibleTiles.x \ 2 To visibleTiles.x \ 2 - 1
            tileZ = tileScreenInfo(idx, jdx).z

            If tileZ <= nearestZ Then
                quad_.coords(0).x = tileScreenInfo(idx + 0, jdx + 0).x
                quad_.coords(0).y = tileScreenInfo(idx + 0, jdx + 0).y

                quad_.coords(1).x = tileScreenInfo(idx + 0, jdx + 1).x
                quad_.coords(1).y = tileScreenInfo(idx + 0, jdx + 1).y

                quad_.coords(2).x = tileScreenInfo(idx + 1, jdx + 1).x
                quad_.coords(2).y = tileScreenInfo(idx + 1, jdx + 1).y

                quad_.coords(3).x = tileScreenInfo(idx + 1, jdx + 0).x
                quad_.coords(3).y = tileScreenInfo(idx + 1, jdx + 0).y

                ' We've got a match for our mouse coords
                If inQuad(screenCoord, quad_) Then
                    outMousePos.x = playerPos.p.x + (world_coord(j) - (playerPos.p.x Mod TILE_WIDTH))
                    outMousePos.y = playerPos.p.z + (world_coord(i) - (playerPos.p.z Mod TILE_HEIGHT))
                    piq = positionInQuad(screenCoord, quad_)
                    outMousePos.x = outMousePos.x + piq.x
                    outMousePos.y = outMousePos.y + piq.y
                    If outMousePos.x < 0 Then
                        outMousePos.x = 0
                    ElseIf outMousePos.x > world_coord(mapWidth - 1) Then
                        outMousePos.x = world_coord(mapWidth - 1)
                    End If
                    If outMousePos.y < 0 Then
                        outMousePos.y = 0
                    ElseIf outMousePos.y > world_coord(mapHeight - 1) Then
                        outMousePos.y = world_coord(mapHeight - 1)
                    End If
                    tileX = map_coord(outMousePos.x)
                    tileY = map_coord(outMousePos.y)
                    ' Store away z value
                    nearestZ = tileZ
                End If
            End If
            jdx = jdx + 1
        Next
        idx = idx + 1
    Next
    worldCoord = outMousePos
End Sub

'*
' * Find the tile the mouse is currently over
' * \todo This is slow - speed it up
Private Sub locateMouse()
    Dim pt As Vector2i
    pt.x = mouseX()
    pt.y = mouseY()
    screenCoordToWorld pt, mousePos, mouseTileX, mouseTileY
End Sub

' (VB) std::remainder(x, y): x - n * y, n the nearest integer to x / y (ties to even)
Private Function remainderf(ByVal x As Single, ByVal y As Single) As Single
    Dim q As Double, n As Double
    q = CDbl(x) / CDbl(y)
    n = Int(q + 0.5)
    If n - q = 0.5 Then
        If (n / 2) <> Int(n / 2) Then n = n - 1    ' a tie: to the even integer
    End If
    remainderf = CSng(CDbl(x) - n * CDbl(y))
End Function

' Render the sky and surroundings
Private Sub renderSurroundings(projectionMatrix As mat4, skyboxViewMatrix As mat4)
    Dim m As mat4
    ' Render skybox relative to ground (i.e. undo player y translation)
    ' then move it somewhat below ground level for the blending effect
    ' rotate it

    If Not gamePaused() Then
        wind = remainderf(wind + graphicsTimeAdjustedIncrement(windSpeed), 360!)
    End If

    ' TODO: skybox needs to be just below lowest point on map (because we have a bottom cap now). Hardcoding for now.
    m = glm_mat4_mul(skyboxViewMatrix, glm_translate(glm_vec3(0!, -500!, 0!)))
    m = glm_mat4_mul(m, glm_rotate(RADIANS(wind), glm_vec3(0!, 1!, 0!)))
    pie_DrawSkybox skybox_scale, VarPtr(projectionMatrix), VarPtr(m)
End Sub

Private Function calculateCameraHeight(ByVal mapHeight_ As Long) As Long
    ' static_cast<int>(std::ceil(static_cast<float>(_mapHeight) / static_cast<float>(HEIGHT_TRACK_INCREMENTS))) * HEIGHT_TRACK_INCREMENTS + CAMERA_PIVOT_HEIGHT
    calculateCameraHeight = CLng(-Int(-CSng(CSng(mapHeight_) / CSng(HEIGHT_TRACK_INCREMENTS)))) * HEIGHT_TRACK_INCREMENTS + CAMERA_PIVOT_HEIGHT
End Function

Public Function calculateCameraHeightAt(ByVal tileX As Long, ByVal tileY As Long) As Long
    calculateCameraHeightAt = calculateCameraHeight(calcAverageTerrainHeight(tileX, tileY))
End Function

' lib/framework/math_ext.h: static inline void solveDifferential2ndOrder(float *y_, float *dydt_, double acceleration,
' double friction, double dt): finds change in y and y' after time dt, according to the differential equation
' y'' = -ay - fy' (VB: std::complex<double> written out as real and imaginary parts)
Private Sub solveDifferential2ndOrder(y_ As Single, dydt_ As Single, ByVal acceleration As Double, ByVal friction As Double, ByVal dt As Double)
    Dim y As Double, dydt As Double, d As Double, sqdR As Double, sqdI As Double
    Dim h1R As Double, h1I As Double, h2R As Double, h2I As Double, e1R As Double, e1I As Double, e2R As Double, e2I As Double
    Dim nR As Double, nI As Double, dR As Double, dI As Double, den As Double, g1R As Double, g1I As Double, g2R As Double, g2I As Double
    Dim aR As Double, aI As Double, bR As Double, bI As Double, ex As Double
    y = y_
    dydt = dydt_
    ' Solution is y = g_1 exp(h_1 t) + g_2 exp(h_2 t), where h_i are the solutions to h^2 + f h + a = 0, which are -f/2 +- sqrt(f^2/4 - a).
    ' At t = 0, g_1 = (y' - h_2 y) / (h_1 - h_2) and g_2 = (y' - h_1 y) / (h_2 - h_1).

    d = friction * friction / 4 - acceleration          ' std::complex<double> d (real)

    If d >= 0 Then                                      ' std::sqrt(d)
        sqdR = Sqr(d)
        sqdI = 0
    Else
        sqdR = 0
        sqdI = Sqr(-d)
    End If
    h1R = -friction / 2 + sqdR
    h1I = sqdI
    h2R = -friction / 2 - sqdR
    h2I = -sqdI
    ex = Exp(h1R * dt)                                  ' e1 = std::exp(h1 * dt)
    e1R = ex * Cos(h1I * dt)
    e1I = ex * Sin(h1I * dt)
    ex = Exp(h2R * dt)                                  ' e2 = std::exp(h2 * dt)
    e2R = ex * Cos(h2I * dt)
    e2I = ex * Sin(h2I * dt)
    ' g1 = (dydt - h2 * y) / (h1 - h2)
    nR = dydt - h2R * y
    nI = -h2I * y
    dR = h1R - h2R
    dI = h1I - h2I
    den = dR * dR + dI * dI
    g1R = (nR * dR + nI * dI) / den
    g1I = (nI * dR - nR * dI) / den
    ' g2 = (dydt - h1 * y) / (h2 - h1)
    nR = dydt - h1R * y
    nI = -h1I * y
    dR = h2R - h1R
    dI = h2I - h1I
    g2R = (nR * dR + nI * dI) / den
    g2I = (nI * dR - nR * dI) / den
    ' *y_ = (g1 * e1 + g2 * e2).real()
    y_ = CSng((g1R * e1R - g1I * e1I) + (g2R * e2R - g2I * e2I)) ' .imag() should be 0.
    ' *dydt_ = (g1 * h1 * e1 + g2 * h2 * e2).real()
    aR = g1R * h1R - g1I * h1I
    aI = g1R * h1I + g1I * h1R
    bR = g2R * h2R - g2I * h2I
    bI = g2R * h2I + g2I * h2R
    dydt_ = CSng((aR * e1R - aI * e1I) + (bR * e2R - bI * e2I)) ' .imag() should be 0.
End Sub

' Smoothly adjust player height to match the desired height
Private Sub trackHeight(ByVal desiredHeight As Long)
    Const minTrackHeightInterval As Long = GAME_TICKS_PER_SEC \ 60
    Dim deltaTrackHeightRealTime As Long, separation As Single

    deltaTrackHeightRealTime = FromU32(U32(realTime) - U32(trackHeight_lastHeightAdjustmentRealTime))
    If ULT(deltaTrackHeightRealTime, minTrackHeightInterval) Then
        ' avoid processing this too rapidly, such as when vsync is disabled
        Exit Sub
    End If
    trackHeight_lastHeightAdjustmentRealTime = realTime

    If (desiredHeight = playerPos.p.y) And ((trackHeight_heightSpeed > -5!) And (trackHeight_heightSpeed < 5!)) Then
        trackHeight_heightSpeed = 0!
        Exit Sub
    End If

    separation = desiredHeight - playerPos.p.y   ' How far are we from desired height?

    ' d^2/dt^2 player.p.y = -ACCEL_CONSTANT * (player.p.y - desiredHeight) - VELOCITY_CONSTANT * d/dt player.p.y
    solveDifferential2ndOrder separation, trackHeight_heightSpeed, ACCEL_CONSTANT, VELOCITY_CONSTANT, CSng(U32(deltaTrackHeightRealTime)) / CSng(GAME_TICKS_PER_SEC)

    ' Adjust the height accordingly
    playerPos.p.y = desiredHeight - Fix(separation)
End Sub
