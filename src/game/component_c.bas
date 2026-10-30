Attribute VB_Name = "component_c"
' src/component.cpp, component.h: the players' colours, the components' and
' droids' display (the interface's buttons and the world's droids:
' displayComponentObject), the droids' destruction effects, whole.
' An iIMDShape * (a display model) and an iIMDBaseShape * are WzGfx's handles
' (0 = nullptr), their fields GfxApi's iIMDShape_* / iIMDBaseShape_*
' accessors; a mat4 goes to WzGfx by address (VarPtr). A BASE_STATS * is a
' stats handle (StatType + index), a typed stats pointer its index (-1 =
' nullptr). component.h's IMD macros (BODY_IMD, WEAPON_IMD, ...) are written
' out. displayComponentButtonTemplate's local DROID is a slot of the droid
' pool (DroidContainer_alloc), erased again after the drawing.
Option Explicit

Public Const BLIP_ANIM_DURATION As Long = 200
Public Const OBJECT_RADIUS As Long = 128
Public Const COMPONENT_RADIUS As Long = 64
Public Const DESIGN_DROID_SCALE As Long = 200
Public Const DESIGN_COMPONENT_SCALE As Long = 150
Public Const RESEARCH_COMPONENT_SCALE As Long = 300
Public Const COMP_BUT_SCALE As Long = 100
Public Const DROID_BUT_SCALE As Long = 72
Public Const SMALL_STRUCT_SCALE As Long = 55
Public Const MED_STRUCT_SCALE As Long = 25    'reduced from 30 to fit command centre in window
Public Const LARGE_STRUCT_SCALE As Long = 25
Public Const ULTRA_SMALL_FEATURE_SCALE As Long = 146
Public Const REALLY_SMALL_FEATURE_SCALE As Long = 116
Public Const SMALL_FEATURE_SCALE As Long = 55
Public Const MED_FEATURE_SCALE As Long = 26
Public Const LARGE_FEATURE_SCALE As Long = 16

Public Const TOWER_HEIGHT As Long = 100

Private Const DEFAULT_COMPONENT_TRANSLUCENCY As Long = 128

'VTOL weapon connector start
Private Const VTOL_CONNECTOR_START As Long = 5

' lib/ivis_opengl/pietypes.h's render style flags
Private Const pie_ECM As Long = &H1
Private Const pie_TRANSLUCENT As Long = &H2
Private Const pie_ADDITIVE As Long = &H4
Private Const pie_BUTTON As Long = &H40
Private Const pie_SHADOW As Long = &H80
Private Const pie_FORCELIGHT As Long = &H800
Private Const pie_SHIELD As Long = &H1000

' (display3d.h's BASE_MUZZLE_FLASH_DURATION, EFFECT_MUZZLE_ADDITIVE: display3d_c's)

Private leftFirst As Boolean

' Colour Lookups
' use col = MAX_PLAYERS for anycolour (see multiint.c)
Public Function setPlayerColour(ByVal player As Long, ByVal col As Long) As Boolean
    If Not ULT(player, MAX_PLAYERS) Then
        NetPlay_players_colour(player) = MAX_PLAYERS
        setPlayerColour = True
        Exit Function
    End If
    ' Allow color values from 0 to 15
    If Not ULT(col, 16) Then ASSERT_FAILURE "col < 16", "src/component.cpp:67", "setPlayerColour", "Bad colour setting": setPlayerColour = False: Exit Function
    NetPlay_players_colour(player) = col
    setPlayerColour = True
End Function

' UBYTE getPlayerColour(UDWORD pl)
Public Function getPlayerColour(ByVal pl As Long) As Long
    If Not ULT(pl, MAX_PLAYERS) Then
        getPlayerColour = 0 ' baba, or spectator slot
        Exit Function
    End If
    getPlayerColour = NetPlay_players_colour(pl) And &HFF&
End Function

' (VB) glm::vec3(v.xzy()) of a connector, read through GfxApi
Private Function connectorXZY(ByVal shape As Long, ByVal i As Long) As Vector3f
    Dim x As Long, y As Long, z As Long
    iIMDShape_connector shape, i, x, y, z
    connectorXZY = glm_vec3(CSng(x), CSng(z), CSng(y))
End Function

' (VB) a connector's z
Private Function connectorZ(ByVal shape As Long, ByVal i As Long) As Long
    Dim x As Long, y As Long, z As Long
    iIMDShape_connector shape, i, x, y, z
    connectorZ = z
End Function

' std::round(float): half away from zero
Private Function roundf(ByVal f As Single) As Single
    If f >= 0! Then roundf = Int(f + 0.5!) Else roundf = -Int(-f + 0.5!)
End Function

' (VB) safeGetDisplayModelFromBase (ivisdef.h): a base model's display model, nullptr for nullptr
Private Function displayModelOf(ByVal baseShape As Long) As Long
    If baseShape <> 0 Then displayModelOf = iIMDBaseShape_displayModel(baseShape)
End Function

Private Function setMatrix(Position As Vector3i, pRotation As Vector3i, ByVal scale_ As Long) As mat4
    Dim s As Single
    s = scale_ / 100!
    setMatrix = glm_mat4_mul(glm_mat4_mul(glm_mat4_mul(glm_mat4_mul( _
        glm_translate(glm_vec3(CSng(Position.x), CSng(Position.y), CSng(Position.z))), _
        glm_rotate(UNDEG(DEG(pRotation.x)), glm_vec3(1!, 0!, 0!))), _
        glm_rotate(UNDEG(DEG(pRotation.y)), glm_vec3(0!, 1!, 0!))), _
        glm_rotate(UNDEG(DEG(pRotation.z)), glm_vec3(0!, 0!, 1!))), _
        glm_scale(glm_vec3(s, s, s)))
End Function

Public Function getComponentDroidRadius(ByVal psDroid As Long) As Long
    getComponentDroidRadius = 100
End Function


Public Function getComponentDroidTemplateRadius(ByVal psDroid As Long) As Long
    getComponentDroidTemplateRadius = 100
End Function


Public Function getComponentRadius(ByVal psComponent As Long) As Long ' DISPLAY ONLY
    Dim ComponentIMD As Long, MountIMD As Long, compID As Long
    ComponentIMD = 0
    MountIMD = 0

    compID = StatIsComponent(psComponent)
    If compID >= 0 Then
        StatGetComponentIMD psComponent, compID, ComponentIMD, MountIMD
        If ComponentIMD <> 0 Then
            getComponentRadius = iIMDShape_sradius(ComponentIMD)
            Exit Function
        End If
    End If

    ' VTOL bombs are only stats allowed to have NULL ComponentIMD
    If StatIsComponent(psComponent) <> COMP_WEAPON Then
        If Not (ComponentIMD <> 0) Then ASSERT_FAILURE "ComponentIMD", "src/component.cpp:123", "getComponentRadius", "No ComponentIMD!"
    ElseIf asWeaponStats(psComponent And &HFFFF&).weaponSubClass <> WSC_BOMB _
           And asWeaponStats(psComponent And &HFFFF&).weaponSubClass <> WSC_EMP Then
        If Not (ComponentIMD <> 0) Then ASSERT_FAILURE "ComponentIMD", "src/component.cpp:123", "getComponentRadius", "No ComponentIMD!"
    End If

    getComponentRadius = COMPONENT_RADIUS
End Function


Public Function getResearchRadius(ByVal Stat As Long) As Long ' DISPLAY ONLY
    Dim ResearchIMD As Long
    ResearchIMD = displayModelOf(asResearch(Stat And &HFFFF&).pIMD)

    If ResearchIMD <> 0 Then
        getResearchRadius = iIMDShape_sradius(ResearchIMD)
        Exit Function
    End If

    debug_ 139, LOG_ERROR, "getResearchRadius", "ResearchPIE == NULL"

    getResearchRadius = 100
End Function


Public Function getStructureSizeMax(ByVal psStructure As Long) As Long ' DISPLAY ONLY
    Dim s As Long
    'radius based on base plate size
    s = aStruct(psStructure).pStructureType
    If ULT(asStructureStats(s).baseWidth, asStructureStats(s).baseBreadth) Then
        getStructureSizeMax = asStructureStats(s).baseBreadth
    Else
        getStructureSizeMax = asStructureStats(s).baseWidth
    End If
End Function

Public Function getStructureStatSizeMax(ByVal Stats As Long) As Long ' DISPLAY ONLY
    'radius based on base plate size
    If ULT(asStructureStats(Stats).baseWidth, asStructureStats(Stats).baseBreadth) Then
        getStructureStatSizeMax = asStructureStats(Stats).baseBreadth
    Else
        getStructureStatSizeMax = asStructureStats(Stats).baseWidth
    End If
End Function

Public Function getStructureStatHeight(ByVal psStat As Long) As Long ' DISPLAY ONLY
    Dim pDisplayIMD As Long, mnx As Long, mny As Long, mnz As Long, mxx As Long, mxy As Long, mxz As Long
    If asStructureStats(psStat).pIMD(0) <> 0 Then
        pDisplayIMD = iIMDBaseShape_displayModel(asStructureStats(psStat).pIMD(0))
        iIMDShape_min pDisplayIMD, mnx, mny, mnz
        iIMDShape_max pDisplayIMD, mxx, mxy, mxz
        getStructureStatHeight = (mxy - mny)
        Exit Function
    End If

    getStructureStatHeight = 0
End Function

Private Sub draw_player_3d_shape(ByVal player_index As Long, ByVal shape As Long, modelMatrix As mat4)
    Dim team As Long, imd As Long, teamcolour As Long, identity As mat4
    team = getPlayerColour(player_index)
    identity = glm_mat4(1!)
    imd = getFactionDisplayIMD(getPlayerFaction(player_index), shape)
    Do While imd <> 0
        teamcolour = iIMDShape_getTeamColourForModel(imd, team)
        pie_Draw3DButton imd, teamcolour, VarPtr(modelMatrix), VarPtr(identity)
        imd = iIMDShape_next(imd)
    Loop
End Sub

Public Sub displayIMDButton(ByVal IMDShape As Long, pRotation As Vector3i, Position As Vector3i, ByVal scale_ As Long)
    Dim m As mat4
    m = setMatrix(Position, pRotation, scale_)
    draw_player_3d_shape selectedPlayer, IMDShape, m
End Sub

Private Sub sharedStructureButton(ByVal Stats As Long, ByVal strBaseImd As Long, pRotation As Vector3i, Position As Vector3i, ByVal scale_ As Long)
    Dim baseImd As Long, mountImd(0 To MAX_WEAPONS - 1) As Long, weaponImd(0 To MAX_WEAPONS - 1) As Long
    Dim pos As Vector3i, strImd As Long, matrix As mat4, localMatrix As mat4, i As Long, n As Long
    Dim displayModel As Long, m As mat4
    If strBaseImd = 0 Then Exit Sub

    pos = Position

    strImd = iIMDBaseShape_displayModel(strBaseImd)

    If strImd = 0 Then
        Exit Sub
    End If

    ' HACK HACK HACK!
    ' if its a 'tall thin (ie tower)' structure stat with something on the top - offset the position to show the object on top
    If iIMDShape_connectors_size(strImd) <> 0 And scale_ = SMALL_STRUCT_SCALE Then
        If ULT(TOWER_HEIGHT, getStructureStatHeight(Stats)) Then
            pos.y = pos.y - 20
        End If
    End If

    matrix = setMatrix(pos, pRotation, scale_)

    ' Draw the building's base first
    baseImd = asStructureStats(Stats).pBaseIMD

    If baseImd <> 0 Then
        draw_player_3d_shape selectedPlayer, iIMDBaseShape_displayModel(baseImd), matrix
    End If
    draw_player_3d_shape selectedPlayer, strImd, matrix

    'and draw the turret
    If iIMDShape_connectors_size(strImd) <> 0 Then
        weaponImd(0) = 0
        mountImd(0) = 0
        For i = 0 To asStructureStats(Stats).numWeaps - 1
            weaponImd(i) = 0 'weapon is gun ecm or sensor
            mountImd(i) = 0
        Next
        'get an imd to draw on the connector priority is weapon, ECM, sensor
        'check for weapon
        'can only have the MAX_WEAPONS
        If asStructureStats(Stats).numWeaps > 1 Then n = asStructureStats(Stats).numWeaps Else n = 1
        For i = 0 To n - 1
            'can only have the one
            If asStructureStats(Stats).psWeapStat(i) >= 0 Then
                weaponImd(i) = asWeaponStats(asStructureStats(Stats).psWeapStat(i)).pIMD
                mountImd(i) = asWeaponStats(asStructureStats(Stats).psWeapStat(i)).pMountGraphic
            End If

            If weaponImd(i) = 0 Then
                'check for ECM
                If asStructureStats(Stats).pECM >= 0 Then
                    weaponImd(i) = asECMStats(asStructureStats(Stats).pECM).pIMD
                    mountImd(i) = asECMStats(asStructureStats(Stats).pECM).pMountGraphic
                End If
            End If

            If weaponImd(i) = 0 Then
                'check for sensor
                If asStructureStats(Stats).pSensor >= 0 Then
                    weaponImd(i) = asSensorStats(asStructureStats(Stats).pSensor).pIMD
                    mountImd(i) = asSensorStats(asStructureStats(Stats).pSensor).pMountGraphic
                End If
            End If
        Next

        'draw Weapon/ECM/Sensor for structure
        If weaponImd(0) <> 0 Then
            For i = 0 To n - 1
                If Not ULT(i, iIMDShape_connectors_size(strImd)) Then
                    Exit For
                End If
                localMatrix = glm_translate(connectorXZY(strImd, i))
                If mountImd(i) <> 0 Then
                    displayModel = iIMDBaseShape_displayModel(mountImd(i))
                    m = glm_mat4_mul(matrix, localMatrix)
                    draw_player_3d_shape selectedPlayer, displayModel, m
                    If iIMDShape_connectors_size(displayModel) <> 0 Then
                        localMatrix = glm_mat4_mul(localMatrix, glm_translate(connectorXZY(displayModel, 0)))
                    End If
                End If
                m = glm_mat4_mul(matrix, localMatrix)
                draw_player_3d_shape selectedPlayer, iIMDBaseShape_displayModel(weaponImd(i)), m
                'we have a droid weapon so do we draw a muzzle flash
            Next
        End If
    End If
End Sub

Public Sub displayStructureButton(ByVal psStructure As Long, rotationV As Vector3i, Position As Vector3i, ByVal scale_ As Long)
    sharedStructureButton aStruct(psStructure).pStructureType, aStruct(psStructure).sDisplay.imd, rotationV, Position, scale_
End Sub

' Stats: a STRUCTURE_STATS * (its index, -1 = nullptr)
Public Sub displayStructureStatButton(ByVal Stats As Long, rotationV As Vector3i, Position As Vector3i, ByVal scale_ As Long)
    If Not (Stats >= 0) Then ASSERT_FAILURE "Stats != nullptr", "src/component.cpp:291", "displayStructureStatButton", "Stats is NULL": Exit Sub
    If asStructureStats(Stats).pIMD_size = 0 Then
        Exit Sub ' error with loaded stats/models - probably due to a mod
    End If
    If asStructureStats(Stats).pIMD(0) = 0 Then
        Exit Sub ' error with loaded stats/models - probably due to a mod
    End If
    sharedStructureButton Stats, asStructureStats(Stats).pIMD(0), rotationV, Position, scale_
End Sub

' Render a component given a BASE_STATS structure.
'
Public Sub displayComponentButton(ByVal Stat As Long, pRotation As Vector3i, Position As Vector3i, ByVal scale_ As Long)
    Dim ComponentIMD As Long, MountIMD As Long, compID As Long, matrix As mat4
    ComponentIMD = 0
    MountIMD = 0
    compID = StatIsComponent(Stat)

    If compID >= 0 Then
        StatGetComponentIMD Stat, compID, ComponentIMD, MountIMD
    Else
        Exit Sub
    End If
    matrix = setMatrix(Position, pRotation, scale_)

    ' VTOL bombs are only stats allowed to have NULL ComponentIMD
    If StatIsComponent(Stat) <> COMP_WEAPON Then
        If Not (ComponentIMD <> 0) Then ASSERT_FAILURE "ComponentIMD", "src/component.cpp:322", "displayComponentButton", "No ComponentIMD"
    ElseIf asWeaponStats(Stat And &HFFFF&).weaponSubClass <> WSC_BOMB _
           And asWeaponStats(Stat And &HFFFF&).weaponSubClass <> WSC_EMP Then
        If Not (ComponentIMD <> 0) Then ASSERT_FAILURE "ComponentIMD", "src/component.cpp:322", "displayComponentButton", "No ComponentIMD"
    End If

    If MountIMD <> 0 Then
        draw_player_3d_shape selectedPlayer, MountIMD, matrix

        ' translate for weapon mount point
        If iIMDShape_connectors_size(MountIMD) <> 0 Then
            matrix = glm_mat4_mul(matrix, glm_translate(connectorXZY(MountIMD, 0)))
        End If
    End If
    If ComponentIMD <> 0 Then
        draw_player_3d_shape selectedPlayer, ComponentIMD, matrix
    End If
End Sub


' Render a research item given a BASE_STATS structure.
'
Public Sub displayResearchButton(ByVal Stat As Long, pRotation As Vector3i, Position As Vector3i, ByVal scale_ As Long)
    Dim ResearchIMD As Long, MountIMD As Long, matrix As mat4
    If Not (Stat <> 0) Then ASSERT_FAILURE "Stat != nullptr", "src/component.cpp:346", "displayResearchButton", "Stat is NULL": Exit Sub
    ResearchIMD = asResearch(Stat And &HFFFF&).pIMD
    MountIMD = asResearch(Stat And &HFFFF&).pIMD2

    If Not (ResearchIMD <> 0) Then ASSERT_FAILURE "ResearchIMD", "src/component.cpp:350", "displayResearchButton", "ResearchIMD is NULL"
    If ResearchIMD <> 0 Then
        matrix = setMatrix(Position, pRotation, scale_)

        If MountIMD <> 0 Then
            draw_player_3d_shape selectedPlayer, iIMDBaseShape_displayModel(MountIMD), matrix
        End If
        draw_player_3d_shape selectedPlayer, iIMDBaseShape_displayModel(ResearchIMD), matrix
    End If
End Sub


Private Function getLeftPropulsionIMD(ByVal psDroid As Long) As Long
    Dim bodyStat As Long, propStat As Long
    bodyStat = aDroid(psDroid).asBits(COMP_BODY)
    propStat = aDroid(psDroid).asBits(COMP_PROPULSION)
    getLeftPropulsionIMD = asBodyStats(bodyStat).ppIMDList(propStat * NUM_PROP_SIDES + LEFT_PROP)
End Function

Private Function getRightPropulsionIMD(ByVal psDroid As Long) As Long
    Dim bodyStat As Long, propStat As Long
    bodyStat = aDroid(psDroid).asBits(COMP_BODY)
    propStat = aDroid(psDroid).asBits(COMP_PROPULSION)
    getRightPropulsionIMD = asBodyStats(bodyStat).ppIMDList(propStat * NUM_PROP_SIDES + RIGHT_PROP)
End Function

' modelMatrix: C++ takes it by value (a copy here)
Public Sub drawMuzzleFlash(sWeap As WEAPON, ByVal weaponImd As Long, ByVal flashImd As Long, ByVal buildingBrightness As Long, ByVal pieFlag As Long, ByVal iPieData As Long, modelMatrix_ As mat4, viewMatrix As mat4, ByVal heightAboveTerrain As Single, Optional ByVal colour As Long = 0)
    Dim connector_num As Long, modelMatrix As mat4, nConnectors As Long, DEFAULT_ANIM_INTERVAL As Long
    Dim animRate As Long, frame As Long
    If weaponImd = 0 Or flashImd = 0 Then Exit Sub
    If iIMDShape_connectors_size(weaponImd) = 0 Or ULT(graphicsTime, sWeap.lastFired) Then
        Exit Sub
    End If
    modelMatrix = modelMatrix_

    connector_num = 0
    nConnectors = iIMDShape_connectors_size(weaponImd)

    ' which barrel is firing if model have multiple muzzle connectors?
    If sWeap.shotsFired <> 0 And (nConnectors > 1) Then
        ' shoot first, draw later - substract one shot to get correct results
        connector_num = FromU32(U32(sWeap.shotsFired - 1) - Int(U32(sWeap.shotsFired - 1) / nConnectors) * nConnectors)
    End If

    ' Now we need to move to the end of the firing barrel
    modelMatrix = glm_mat4_mul(modelMatrix, glm_translate(connectorXZY(weaponImd, connector_num)))
    heightAboveTerrain = heightAboveTerrain + connectorZ(weaponImd, connector_num)

    ' assume no clan colours for muzzle effects
    If iIMDShape_numFrames(flashImd) = 0 Or iIMDShape_animInterval(flashImd) <= 0 Then
        ' no anim so display one frame for a fixed time
        If Not ULT(graphicsTime, sWeap.lastFired) And ULT(graphicsTime, sWeap.lastFired + BASE_MUZZLE_FLASH_DURATION) Then
            pie_Draw3DShape flashImd, 0, colour, buildingBrightness, pieFlag Or pie_ADDITIVE, EFFECT_MUZZLE_ADDITIVE, VarPtr(modelMatrix), VarPtr(viewMatrix), -heightAboveTerrain
        End If
    ElseIf Not ULT(graphicsTime, sWeap.lastFired) Then
        ' animated muzzle
        DEFAULT_ANIM_INTERVAL = 17 ' A lot of PIE files specify 1, which is too small, so set something bigger as a fallback
        animRate = iIMDShape_animInterval(flashImd)
        If animRate < DEFAULT_ANIM_INTERVAL Then animRate = DEFAULT_ANIM_INTERVAL
        frame = FromU32(Int(U32(graphicsTime - sWeap.lastFired) / animRate))
        If frame < iIMDShape_numFrames(flashImd) Then
            pie_Draw3DShape flashImd, frame, colour, buildingBrightness, pieFlag Or pie_ADDITIVE, EFFECT_MUZZLE_ADDITIVE, VarPtr(modelMatrix), VarPtr(viewMatrix), -heightAboveTerrain
        End If
    End If
End Sub

' Assumes matrix context is already set
' this is able to handle multiple weapon graphics now
' removed mountRotation,they get such stuff from psObj directly now
Private Function displayCompObj(ByVal psDroid As Long, ByVal bButton As Boolean, modelMatrix2 As mat4, viewMatrix As mat4) As Boolean
    Dim psMoveAnim As Long, psStillAnim As Long, iConnector As Long, psPropStats As Long
    Dim pieFlag As Long, iPieData As Long, shieldPieFlag As Long, iShieldPieData As Long
    Dim brightness As Long, colour As Long, i As Long, didDrawSomething As Boolean
    Dim factor As Single, modifiedModelMatrix As mat4, psTile As Long, psShapeProp As Long
    Dim psBaseShapeBody As Long, psShapeBody As Long, strImd As Long, bodyStats As Long, displayModel As Long
    Dim rot As Rotation, localModelMatrix As mat4, localHeightAboveTerrain As Single, bodyConnectorIndex As Long
    Dim psBaseShape As Long, psShape As Long, recoilValue As Long, flashBaseModel As Long, pMuzzleFlash As Long
    Dim psBaseMountShape As Long, psMountShape As Long, st As Spacetime, propIdx As Long, wStat As Long
    Dim heightAboveMap As Single

    shieldPieFlag = 0
    iShieldPieData = 0
    i = 0
    didDrawSomething = False
    heightAboveMap = aDroid(psDroid).heightAboveMap

    If Not bButton And aDroid(psDroid).shieldPoints > 0 Then
        If droidGetMaxShieldPoints(psDroid) > 0 Then
            factor = CSng(aDroid(psDroid).shieldPoints) / droidGetMaxShieldPoints(psDroid)
            iShieldPieData = CLng(roundf(255! * factor))
            shieldPieFlag = pie_FORCELIGHT Or pie_TRANSLUCENT Or pie_SHIELD
        End If
    End If

    modifiedModelMatrix = modelMatrix2

    If ULT(aDroid(psDroid).timeLastHit - graphicsTime, ELEC_DAMAGE_DURATION) And aDroid(psDroid).lastHitWeapon = WSC_ELECTRONIC And Not gamePaused() Then
        colour = getPlayerColour(rand() Mod MAX_PLAYERS)
    Else
        colour = getPlayerColour(aDroid(psDroid).player)
    End If

    ' get propulsion stats
    psPropStats = DROID_getPropulsionStats(psDroid)
    If Not (psPropStats >= 0) Then ASSERT_FAILURE "psPropStats != nullptr", "src/component.cpp:455", "displayCompObj", "invalid propulsion stats pointer": displayCompObj = didDrawSomething: Exit Function

    'set pieflag for button object or ingame object
    If bButton Then
        pieFlag = pie_BUTTON
        brightness = psPalette(WZCOL_WHITE)
    Else
        pieFlag = pie_SHADOW
        brightness = pal_SetBrightness(aDroid(psDroid).illumination)
        ' NOTE: Beware of transporters that are offscreen, on a mission!  We should *not* be checking tiles at this point in time!
        If Not DROID_isTransporter(psDroid) And Not missionIsOffworld() Then
            psTile = worldTile(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y)
            If (psMapTiles(psTile).jammerBits And alliancebits(aDroid(psDroid).player)) <> 0 Then
                pieFlag = pieFlag Or pie_ECM
            End If
        End If
    End If

    ' set default components transparent
    If aDroid(psDroid).asBits(COMP_PROPULSION) = 0 Then
        pieFlag = pieFlag Or pie_TRANSLUCENT
        iPieData = DEFAULT_COMPONENT_TRANSLUCENCY
    Else
        iPieData = 0
    End If

    If Not bButton And asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_PROPELLOR Then
        ' FIXME: change when adding submarines to the game
        modifiedModelMatrix = glm_mat4_mul(modifiedModelMatrix, glm_translate(glm_vec3(0!, -world_coord(1) / 2.3!, 0!)))
    End If

    If leftFirst Then psShapeProp = getLeftPropulsionIMD(psDroid) Else psShapeProp = getRightPropulsionIMD(psDroid)
    If psShapeProp <> 0 Then
        If pie_Draw3DShape(iIMDBaseShape_displayModel(psShapeProp), 0, colour, brightness, pieFlag, iPieData, VarPtr(modifiedModelMatrix), VarPtr(viewMatrix), -heightAboveMap) Then
            didDrawSomething = True
        End If
        If Not bButton And aDroid(psDroid).shieldPoints > 0 Then
            If pie_Draw3DShape(iIMDBaseShape_displayModel(psShapeProp), 0, colour, brightness, shieldPieFlag, iShieldPieData, VarPtr(modifiedModelMatrix), VarPtr(viewMatrix), -heightAboveMap) Then
                didDrawSomething = True
            End If
        End If
    End If

    ' set default components transparent
    If aDroid(psDroid).asBits(COMP_BODY) = 0 Then
        pieFlag = pieFlag Or pie_TRANSLUCENT
        iPieData = DEFAULT_COMPONENT_TRANSLUCENCY
    Else
        pieFlag = pieFlag And Not pie_TRANSLUCENT
        iPieData = 0
    End If

    ' Get the body graphic now
    psBaseShapeBody = asBodyStats(DROID_getBodyStats(psDroid)).pIMD     ' BODY_IMD(psDroid, psDroid->player)
    psShapeBody = displayModelOf(psBaseShapeBody)
    If psShapeBody <> 0 Then
        strImd = psShapeBody
        If aDroid(psDroid).droidType = DROID_PERSON Then
            ' previously, used to scale the model by 0.75 - no longer needed with the fixed models though...
        End If
        If iIMDShape_objanimpie(strImd, aDroid(psDroid).animationEvent) <> 0 Then
            strImd = iIMDBaseShape_displayModel(iIMDShape_objanimpie(strImd, aDroid(psDroid).animationEvent))
        End If
        Do While strImd <> 0
            If drawShape(strImd, aDroid(psDroid).timeAnimationStarted, colour, brightness, pieFlag, iPieData, modifiedModelMatrix, viewMatrix, -heightAboveMap) Then
                didDrawSomething = True
            End If
            If Not bButton And aDroid(psDroid).shieldPoints > 0 Then
                If drawShape(strImd, aDroid(psDroid).timeAnimationStarted, colour, brightness, shieldPieFlag, iShieldPieData, modifiedModelMatrix, viewMatrix, -heightAboveMap) Then
                    didDrawSomething = True
                End If
            End If
            strImd = iIMDShape_next(strImd)
        Loop
    End If

    ' Render animation effects based on movement or lack thereof, if any
    bodyStats = DROID_getBodyStats(psDroid)
    propIdx = aDroid(psDroid).asBits(COMP_PROPULSION)
    psMoveAnim = asBodyStats(bodyStats).ppMoveIMDList(propIdx)
    psStillAnim = asBodyStats(bodyStats).ppStillIMDList(propIdx)
    If Not bButton And psMoveAnim <> 0 And aDroid(psDroid).sMove.Status <> MOVEINACTIVE Then
        displayModel = iIMDBaseShape_displayModel(psMoveAnim)
        If pie_Draw3DShape(displayModel, getModularScaledGraphicsTime(iIMDShape_animInterval(displayModel), iIMDShape_numFrames(displayModel)), colour, brightness, pie_ADDITIVE, 200, VarPtr(modifiedModelMatrix), VarPtr(viewMatrix), -heightAboveMap) Then
            didDrawSomething = True
        End If
    ElseIf Not bButton And psStillAnim <> 0 Then ' standing still
        displayModel = iIMDBaseShape_displayModel(psStillAnim)
        If pie_Draw3DShape(displayModel, getModularScaledGraphicsTime(iIMDShape_animInterval(displayModel), iIMDShape_numFrames(displayModel)), colour, brightness, 0, 0, VarPtr(modifiedModelMatrix), VarPtr(viewMatrix), -heightAboveMap) Then
            didDrawSomething = True
        End If
    End If

    ' set default components transparent
    If aDroid(psDroid).asWeaps(0).nStat = 0 And _
       aDroid(psDroid).asBits(COMP_SENSOR) = 0 And _
       aDroid(psDroid).asBits(COMP_ECM) = 0 And _
       aDroid(psDroid).asBits(COMP_BRAIN) = 0 And _
       aDroid(psDroid).asBits(COMP_REPAIRUNIT) = 0 And _
       aDroid(psDroid).asBits(COMP_CONSTRUCT) = 0 Then
        pieFlag = pieFlag Or pie_TRANSLUCENT
        iPieData = DEFAULT_COMPONENT_TRANSLUCENCY
    Else
        pieFlag = pieFlag And Not pie_TRANSLUCENT
        iPieData = 0
    End If

    If psShapeBody <> 0 Then
        If iIMDShape_connectors_size(psShapeBody) <> 0 Then
            ' vtol weapons attach to connector 2 (underneath);
            ' all others to connector 1
            ' VTOL's now skip the first 5 connectors(0 to 4),
            ' VTOL's use 5,6,7,8 etc now
            If asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_LIFT And aDroid(psDroid).droidType = DROID_WEAPON Then
                iConnector = VTOL_CONNECTOR_START
            Else
                iConnector = 0
            End If

            Select Case aDroid(psDroid).droidType
            Case DROID_DEFAULT, DROID_TRANSPORTER, DROID_SUPERTRANSPORTER, DROID_CYBORG, DROID_CYBORG_SUPER, DROID_WEAPON, _
                 DROID_COMMAND     ' command droids have a weapon to store all the graphics
                '   Get the mounting graphic - we've already moved to the right position
                '   Allegedly - all droids will have a mount graphic so this shouldn't
                '   fall on it's arse......
                ' Double check that the weapon droid actually has any
                For i = 0 To aDroid(psDroid).numWeaps - 1
                    If (aDroid(psDroid).asWeaps(i).nStat > 0 Or aDroid(psDroid).droidType = DROID_DEFAULT) _
                       And iIMDShape_connectors_size(psShapeBody) <> 0 Then
                        rot = getInterpolatedWeaponRotation(psDroid, i, graphicsTime)

                        localModelMatrix = modifiedModelMatrix
                        localHeightAboveTerrain = heightAboveMap

                        'to skip number of VTOL_CONNECTOR_START ground unit connectors
                        If iConnector < VTOL_CONNECTOR_START Then bodyConnectorIndex = i Else bodyConnectorIndex = iConnector + i
                        If bodyConnectorIndex < iIMDShape_connectors_size(psShapeBody) Then
                            localModelMatrix = glm_mat4_mul(localModelMatrix, glm_translate(connectorXZY(psShapeBody, bodyConnectorIndex)))
                            localHeightAboveTerrain = localHeightAboveTerrain + connectorZ(psShapeBody, bodyConnectorIndex)
                        Else
                            debug_ 639, LOG_INFO, "displayCompObj", "Model lacks sufficient connectors?: %s", modelName(psShapeBody)
                        End If
                        localModelMatrix = glm_mat4_mul(localModelMatrix, glm_rotate(UNDEG(-rot.direction), glm_vec3(0!, 1!, 0!)))

                        ' vtol weapons inverted
                        If iConnector >= VTOL_CONNECTOR_START Then
                            'this might affect gun rotation
                            localModelMatrix = glm_mat4_mul(localModelMatrix, glm_rotate(UNDEG(65536 \ 2), glm_vec3(0!, 0!, 1!)))
                        End If

                        ' Get the mount graphic
                        wStat = DROID_getWeaponStats(psDroid, i)
                        psBaseShape = asWeaponStats(wStat).pMountGraphic       ' WEAPON_MOUNT_IMD(psDroid, i)
                        psShape = displayModelOf(psBaseShape)

                        recoilValue = getRecoil(aDroid(psDroid).asWeaps(i))
                        localModelMatrix = glm_mat4_mul(localModelMatrix, glm_translate(glm_vec3(0!, 0!, recoilValue / 3!)))

                        ' Draw it
                        If psShape <> 0 Then
                            If pie_Draw3DShape(psShape, 0, colour, brightness, pieFlag, iPieData, VarPtr(localModelMatrix), VarPtr(viewMatrix), -localHeightAboveTerrain) Then
                                didDrawSomething = True
                            End If
                            If Not bButton And aDroid(psDroid).shieldPoints > 0 Then
                                If pie_Draw3DShape(psShape, 0, colour, brightness, shieldPieFlag, iShieldPieData, VarPtr(localModelMatrix), VarPtr(viewMatrix), -localHeightAboveTerrain) Then
                                    didDrawSomething = True
                                End If
                            End If
                        End If
                        localModelMatrix = glm_mat4_mul(localModelMatrix, glm_translate(glm_vec3(0!, 0!, CSng(recoilValue))))

                        ' translate for weapon mount point
                        If psShape <> 0 Then
                            If iIMDShape_connectors_size(psShape) <> 0 Then
                                localModelMatrix = glm_mat4_mul(localModelMatrix, glm_translate(connectorXZY(psShape, 0)))
                                localHeightAboveTerrain = localHeightAboveTerrain + connectorZ(psShape, 0)
                            End If
                        End If

                        ' vtol weapons inverted
                        If iConnector >= VTOL_CONNECTOR_START Then
                            'pitch the barrel down
                            localModelMatrix = glm_mat4_mul(localModelMatrix, glm_rotate(UNDEG(-rot.pitch), glm_vec3(1!, 0!, 0!)))
                        Else
                            'pitch the barrel up
                            localModelMatrix = glm_mat4_mul(localModelMatrix, glm_rotate(UNDEG(rot.pitch), glm_vec3(1!, 0!, 0!)))
                        End If

                        ' Get the weapon (gun?) graphic
                        psBaseShape = asWeaponStats(wStat).pIMD                ' WEAPON_IMD(psDroid, i)
                        psShape = displayModelOf(psBaseShape)

                        ' We have a weapon so we draw it and a muzzle flash from weapon connector
                        If psShape <> 0 Then
                            If pie_Draw3DShape(psShape, 0, colour, brightness, pieFlag, iPieData, VarPtr(localModelMatrix), VarPtr(viewMatrix), -localHeightAboveTerrain) Then
                                didDrawSomething = True
                            End If
                            If Not bButton And aDroid(psDroid).shieldPoints > 0 Then
                                If pie_Draw3DShape(psShape, 0, colour, brightness, shieldPieFlag, iShieldPieData, VarPtr(localModelMatrix), VarPtr(viewMatrix), -localHeightAboveTerrain) Then
                                    didDrawSomething = True
                                End If
                            End If
                            flashBaseModel = asWeaponStats(wStat).pMuzzleGraphic  ' MUZZLE_FLASH_PIE(psDroid, i)
                            pMuzzleFlash = displayModelOf(flashBaseModel)
                            drawMuzzleFlash aDroid(psDroid).asWeaps(i), psShape, pMuzzleFlash, brightness, pieFlag, iPieData, localModelMatrix, viewMatrix, localHeightAboveTerrain
                        End If
                    End If
                Next

            Case DROID_SENSOR, DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT, DROID_ECM, DROID_REPAIR, DROID_CYBORG_REPAIR
                rot = getInterpolatedWeaponRotation(psDroid, 0, graphicsTime)
                psBaseShape = 0
                psBaseMountShape = 0

                Select Case aDroid(psDroid).droidType
                Case DROID_SENSOR
                    psBaseMountShape = asSensorStats(DROID_getSensorStats(psDroid)).pMountGraphic       ' SENSOR_MOUNT_IMD
                    ' Get the sensor graphic, assuming it's there
                    psBaseShape = asSensorStats(DROID_getSensorStats(psDroid)).pIMD                     ' SENSOR_IMD
                Case DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT
                    psBaseMountShape = asConstructStats(DROID_getConstructStats(psDroid)).pMountGraphic ' CONSTRUCT_MOUNT_IMD
                    ' Get the construct graphic assuming it's there
                    psBaseShape = asConstructStats(DROID_getConstructStats(psDroid)).pIMD               ' CONSTRUCT_IMD
                Case DROID_ECM
                    psBaseMountShape = asECMStats(DROID_getECMStats(psDroid)).pMountGraphic             ' ECM_MOUNT_IMD
                    ' Get the ECM graphic assuming it's there....
                    psBaseShape = asECMStats(DROID_getECMStats(psDroid)).pIMD                           ' ECM_IMD
                Case DROID_REPAIR, DROID_CYBORG_REPAIR
                    psBaseMountShape = asRepairStats(DROID_getRepairStats(psDroid)).pMountGraphic       ' REPAIR_MOUNT_IMD
                    ' Get the Repair graphic assuming it's there....
                    psBaseShape = asRepairStats(DROID_getRepairStats(psDroid)).pIMD                     ' REPAIR_IMD
                Case Else
                    ASSERT_FAILURE "false", "src/component.cpp:733", "displayCompObj", "Bad component type"
                End Select

                psShape = displayModelOf(psBaseShape)
                psMountShape = displayModelOf(psBaseMountShape)

                '   Get the mounting graphic - we've already moved to the right position
                '   Allegedly - all droids will have a mount graphic so this shouldn't
                '   fall on it's arse......
                'sensor and cyborg and ecm uses connectors[0]

                localModelMatrix = modifiedModelMatrix
                localHeightAboveTerrain = heightAboveMap
                ' vtol weapons inverted
                If iConnector >= VTOL_CONNECTOR_START Then
                    'this might affect gun rotation
                    localModelMatrix = glm_mat4_mul(localModelMatrix, glm_rotate(UNDEG(65536 \ 2), glm_vec3(0!, 0!, 1!)))
                End If

                localModelMatrix = glm_mat4_mul(localModelMatrix, glm_translate(connectorXZY(psShapeBody, 0)))
                localHeightAboveTerrain = localHeightAboveTerrain + connectorZ(psShapeBody, 0)

                localModelMatrix = glm_mat4_mul(localModelMatrix, glm_rotate(UNDEG(-rot.direction), glm_vec3(0!, 1!, 0!)))
                ' Draw it
                If psMountShape <> 0 Then
                    If pie_Draw3DShape(psMountShape, 0, colour, brightness, pieFlag, iPieData, VarPtr(localModelMatrix), VarPtr(viewMatrix), -localHeightAboveTerrain) Then
                        didDrawSomething = True
                    End If
                End If

                ' translate for construct mount point if cyborg
                If DROID_isCyborg(psDroid) And psMountShape <> 0 Then
                    If iIMDShape_connectors_size(psMountShape) <> 0 Then
                        localModelMatrix = glm_mat4_mul(localModelMatrix, glm_translate(connectorXZY(psMountShape, 0)))
                        localHeightAboveTerrain = localHeightAboveTerrain + connectorZ(psMountShape, 0)
                    End If
                End If

                ' Draw it
                If psShape <> 0 Then
                    If pie_Draw3DShape(psShape, 0, colour, brightness, pieFlag, iPieData, VarPtr(localModelMatrix), VarPtr(viewMatrix), -localHeightAboveTerrain) Then
                        didDrawSomething = True
                    End If

                    ' In repair droid case only:
                    If (aDroid(psDroid).droidType = DROID_REPAIR Or aDroid(psDroid).droidType = DROID_CYBORG_REPAIR) And _
                       iIMDShape_connectors_size(psShape) <> 0 And aDroid(psDroid).action = DACTION_DROIDREPAIR Then
                        st = interpolateObjectSpacetime(psDroid, graphicsTime)
                        localModelMatrix = glm_mat4_mul(localModelMatrix, glm_translate(connectorXZY(psShape, 0)))
                        localHeightAboveTerrain = localHeightAboveTerrain + connectorZ(psShape, 0)
                        localModelMatrix = glm_mat4_mul(localModelMatrix, glm_translate(glm_vec3(0!, -20!, 0!)))

                        psShape = iIMDBaseShape_displayModel(getImdFromIndex(MI_FLAME))

                        ' Rotate for droid
                        localModelMatrix = glm_mat4_mul(localModelMatrix, glm_rotate(UNDEG(st.rot.direction), glm_vec3(0!, 1!, 0!)))
                        localModelMatrix = glm_mat4_mul(localModelMatrix, glm_rotate(UNDEG(-st.rot.pitch), glm_vec3(1!, 0!, 0!)))
                        localModelMatrix = glm_mat4_mul(localModelMatrix, glm_rotate(UNDEG(-st.rot.roll), glm_vec3(0!, 0!, 1!)))
                        'rotate Y
                        localModelMatrix = glm_mat4_mul(localModelMatrix, glm_rotate(UNDEG(rot.direction), glm_vec3(0!, 1!, 0!)))

                        localModelMatrix = glm_mat4_mul(localModelMatrix, glm_rotate(UNDEG(-playerPos.r.y), glm_vec3(0!, 1!, 0!)))
                        localModelMatrix = glm_mat4_mul(localModelMatrix, glm_rotate(UNDEG(-playerPos.r.x), glm_vec3(1!, 0!, 0!)))

                        If pie_Draw3DShape(psShape, getModularScaledGraphicsTime(iIMDShape_animInterval(psShape), iIMDShape_numFrames(psShape)), 0, brightness, pie_ADDITIVE, 140, VarPtr(localModelMatrix), VarPtr(viewMatrix), -localHeightAboveTerrain) Then
                            didDrawSomething = True
                        End If

'                       localModelMatrix *= glm::rotate(UNDEG(playerPos.r.x), glm::vec3(1.f, 0.f, 0.f)); // Not used?
'                       localModelMatrix *= glm::rotate(UNDEG(playerPos.r.y), glm::vec3(0.f, 1.f, 0.f)); // Not used?
                    End If
                End If
            Case DROID_PERSON
                ' no extra mounts for people
            Case Else
                ASSERT_FAILURE "!""invalid droid type""", "src/component.cpp:840", "displayCompObj", "Whoa! Weirdy type of droid found in drawComponentObject!!!"
            End Select
        End If
    End If

    ' set default components transparent
    If aDroid(psDroid).asBits(COMP_PROPULSION) = 0 Then
        pieFlag = pieFlag Or pie_TRANSLUCENT
        iPieData = DEFAULT_COMPONENT_TRANSLUCENCY
    Else
        pieFlag = pieFlag And Not pie_TRANSLUCENT
        iPieData = 0
    End If

    ' now render the other propulsion side
    If leftFirst Then psShapeProp = getRightPropulsionIMD(psDroid) Else psShapeProp = getLeftPropulsionIMD(psDroid)
    If psShapeProp <> 0 Then
        If pie_Draw3DShape(iIMDBaseShape_displayModel(psShapeProp), 0, colour, brightness, pieFlag, iPieData, VarPtr(modifiedModelMatrix), VarPtr(viewMatrix), -heightAboveMap) Then
            didDrawSomething = True
        End If
        If Not bButton And aDroid(psDroid).shieldPoints > 0 Then
            If pie_Draw3DShape(iIMDBaseShape_displayModel(psShapeProp), 0, colour, brightness, shieldPieFlag, iShieldPieData, VarPtr(modifiedModelMatrix), VarPtr(viewMatrix), -heightAboveMap) Then
                didDrawSomething = True
            End If
        End If
    End If

    displayCompObj = didDrawSomething
End Function


' Render a composite droid given a DROID_TEMPLATE structure.
'
Public Sub displayComponentButtonTemplate(ByVal psTemplate As Long, pRotation As Vector3i, Position As Vector3i, ByVal scale_ As Long)
    Dim matrix As mat4, psButtonDroid As Long, identity As mat4, i As Long
    matrix = setMatrix(Position, pRotation, scale_)

    ' Decide how to sort it.
    leftFirst = angleDelta(DEG(pRotation.y)) < 0

    ' (VB) DROID Droid(0, selectedPlayer): a droid pool slot, erased when drawn
    psButtonDroid = DroidContainer_alloc()
    DROID_construct_ psButtonDroid, 0, selectedPlayer
    For i = 0 To DROID_MAXCOMP - 1
        aDroid(psButtonDroid).asBits(i) = 0
    Next
    droidSetBits psTemplate, psButtonDroid

    aDroid(psButtonDroid).pos = Vector3i_(0, 0, 0)
    aDroid(psButtonDroid).rot.direction = 0
    aDroid(psButtonDroid).rot.pitch = 0
    aDroid(psButtonDroid).rot.roll = 0

    'draw multi component object as a button object
    identity = glm_mat4(1!)
    displayCompObj psButtonDroid, True, matrix, identity
    DroidContainer_erase psButtonDroid
End Sub


' Render a composite droid given a DROID structure.
'
Public Sub displayComponentButtonObject(ByVal psDroid As Long, pRotation As Vector3i, Position As Vector3i, ByVal scale_ As Long)
    Dim difference As Long, matrix As mat4, identity As mat4

    matrix = setMatrix(Position, pRotation, scale_)

    ' Decide how to sort it.
    difference = pRotation.y Mod 360

    leftFirst = Not ((difference > 0 And difference < 180) Or difference < -180)

    ' And render the composite object.
    'draw multi component object as a button object
    identity = glm_mat4(1!)
    displayCompObj psDroid, True, matrix, identity
End Sub

' Assumes matrix context is already set
' multiple turrets display removed the pointless mountRotation
Public Sub displayComponentObject(ByVal psDroid As Long, viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim position As Vector3i, rotationV As Vector3i, st As Spacetime, modelMatrix As mat4, pvm As mat4
    Dim effectPosition As Vector3i, frame As Long, identityMVP As mat4, closeDist As Long
    st = interpolateObjectSpacetime(psDroid, graphicsTime)

    leftFirst = angleDelta(playerPos.r.y - st.rot.direction) <= 0

    ' Get the real position
    position.x = st.pos.x
    position.z = -(st.pos.y)
    position.y = st.pos.z

    If DROID_isTransporter(psDroid) Then
        position.y = position.y + bobTransporterHeight()
    End If

    ' Get all the pitch,roll,yaw info
    rotationV.y = -st.rot.direction
    rotationV.x = st.rot.pitch
    rotationV.z = st.rot.roll

    ' Translate origin
    ' Rotate for droid
    modelMatrix = glm_mat4_mul(glm_mat4_mul(glm_mat4_mul( _
        glm_translate(glm_vec3(CSng(position.x), CSng(position.y), CSng(position.z))), _
        glm_rotate(UNDEG(rotationV.y), glm_vec3(0!, 1!, 0!))), _
        glm_rotate(UNDEG(rotationV.x), glm_vec3(1!, 0!, 0!))), _
        glm_rotate(UNDEG(rotationV.z), glm_vec3(0!, 0!, 1!)))

    If ULT(aDroid(psDroid).timeLastHit - graphicsTime, ELEC_DAMAGE_DURATION) And aDroid(psDroid).lastHitWeapon = WSC_ELECTRONIC Then
        modelMatrix = glm_mat4_mul(modelMatrix, objectShimmy(psDroid))
    End If

    ' now check if the projected circle is within the screen boundaries
    pvm = glm_mat4_mul(perspectiveViewMatrix, modelMatrix)
    If getIsCloseDistance() Then closeDist = 150 Else closeDist = 25
    If Not clipDroidOnScreen(psDroid, pvm, closeDist) Then
        Exit Sub
    End If

    If aDroid(psDroid).lastHitWeapon = WSC_EMP And ULT(graphicsTime - aDroid(psDroid).timeLastHit, EMP_DISABLE_TIME) Then
        'add an effect on the droid
        effectPosition.x = st.pos.x + (20 - rand() Mod 40)  ' DROID_EMP_SPREAD
        effectPosition.y = st.pos.z + rand() Mod 8
        effectPosition.z = st.pos.y + (20 - rand() Mod 40)  ' DROID_EMP_SPREAD
        effectGiveAuxVar 90 + rand() Mod 20
        addEffect effectPosition, EFFECT_EXPLOSION, EXPLOSION_TYPE_PLASMA, False, 0, 0, graphicsTime
    End If

    If BASE_OBJECT_visibleForLocalDisplay(psDroid) = UBYTE_MAX Then
        'ingame not button object
        'should render 3 mounted weapons now
        If displayCompObj(psDroid, False, modelMatrix, viewMatrix) Then
            ' did draw something to the screen - update the framenumber
            aDroid(psDroid).sDisplay.frameNumber = frameGetFrameNumber()
        End If

        ' set up all the screen coords stuff
        calcScreenCoords psDroid, pvm
    Else
        frame = FromU32(Int(U32(graphicsTime) / BLIP_ANIM_DURATION) + (U32(aDroid(psDroid).id) - Int(U32(aDroid(psDroid).id) / 8192#) * 8192#)) ' de-sync the blip effect, but don't overflow the int
        If pie_Draw3DShape(iIMDBaseShape_displayModel(getImdFromIndex(MI_BLIP)), frame, 0, psPalette(WZCOL_WHITE), pie_ADDITIVE, BASE_OBJECT_visibleForLocalDisplay(psDroid) \ 2, VarPtr(modelMatrix), VarPtr(viewMatrix)) Then
            aDroid(psDroid).sDisplay.frameNumber = frameGetFrameNumber()
        End If
    End If
End Sub


Public Sub destroyFXDroid(ByVal psDroid As Long, ByVal impactTime As Long, velocity As Vector3f)
    Dim i As Long, psImd As Long, maxHorizontalScatter As Long, heightScatter As Long
    Dim a As Long, r As Long, horizontalScatter As Vector2i, pos As Vector3i
    If vbEffectsOff Then Exit Sub       ' (VB) see effects_c
    For i = 0 To 4
        psImd = 0

        maxHorizontalScatter = TILE_UNITS \ 4
        heightScatter = TILE_UNITS \ 5
        ' iSinCosR(rand(), rand() % maxHorizontalScatter) (the arguments taken left to right)
        a = rand() And &HFFFF&
        r = rand() Mod maxHorizontalScatter
        horizontalScatter.x = iSinR(a, r)
        horizontalScatter.y = iCosR(a, r)

        ' (psDroid->pos + Vector3i(horizontalScatter, 16 + heightScatter)).xzy()
        pos.x = aDroid(psDroid).pos.x + horizontalScatter.x
        pos.y = aDroid(psDroid).pos.z + 16 + heightScatter
        pos.z = aDroid(psDroid).pos.y + horizontalScatter.y
        Select Case i
        Case 0
            Select Case aDroid(psDroid).droidType
            Case DROID_DEFAULT, DROID_CYBORG, DROID_CYBORG_SUPER, DROID_CYBORG_CONSTRUCT, DROID_CYBORG_REPAIR, DROID_WEAPON, DROID_COMMAND
                If aDroid(psDroid).numWeaps > 0 Then
                    If aDroid(psDroid).asWeaps(0).nStat > 0 Then
                        psImd = asWeaponStats(DROID_getWeaponStats(psDroid, 0)).pMountGraphic ' WEAPON_MOUNT_IMD(psDroid, 0)
                    End If
                End If
            Case Else
            End Select
        Case 1
            Select Case aDroid(psDroid).droidType
            Case DROID_DEFAULT, DROID_CYBORG, DROID_CYBORG_SUPER, DROID_CYBORG_CONSTRUCT, DROID_CYBORG_REPAIR, DROID_WEAPON, DROID_COMMAND
                If aDroid(psDroid).numWeaps <> 0 Then
                    ' get main weapon
                    psImd = asWeaponStats(DROID_getWeaponStats(psDroid, 0)).pIMD        ' WEAPON_IMD(psDroid, 0)
                End If
            Case Else
            End Select
        End Select
        If psImd = 0 Then
            psImd = getRandomDebrisImd()
        End If
        ' Tell the effect system that it needs to use this player's color for the next effect
        SetEffectForPlayer aDroid(psDroid).player
        addEffect_velocity pos, EFFECT_GRAVITON, GRAVITON_TYPE_EMITTING_DR, True, iIMDBaseShape_displayModel(psImd), getPlayerColour(aDroid(psDroid).player), impactTime, velocity
    Next
End Sub


Public Sub compPersonToBits(ByVal psDroid As Long, velocity As Vector3f)
    Dim position As Vector3i, headImd As Long, legsImd As Long, armImd As Long, bodyImd As Long, col As Long

    If vbEffectsOff Then Exit Sub       ' (VB) see effects_c
    If BASE_OBJECT_visibleForLocalDisplay(psDroid) = 0 Then ' display only - should not affect game state
        ' We can't see the person or cyborg - so get out
        Exit Sub
    End If
    ' get bits pointers according to whether baba or cyborg
    If DROID_isCyborg(psDroid) Then
        ' This is probably unused now, since there's a more appropriate effect for cyborgs.
        headImd = getImdFromIndex(MI_CYBORG_HEAD)
        legsImd = getImdFromIndex(MI_CYBORG_LEGS)
        armImd = getImdFromIndex(MI_CYBORG_ARM)
        bodyImd = getImdFromIndex(MI_CYBORG_BODY)
    Else
        headImd = getImdFromIndex(MI_BABA_HEAD)
        legsImd = getImdFromIndex(MI_BABA_LEGS)
        armImd = getImdFromIndex(MI_BABA_ARM)
        bodyImd = getImdFromIndex(MI_BABA_BODY)
    End If

    ' Get where he's at
    position.x = aDroid(psDroid).pos.x
    position.y = aDroid(psDroid).pos.z + 1
    position.z = aDroid(psDroid).pos.y

    ' Tell about player colour
    col = getPlayerColour(aDroid(psDroid).player)

    addEffect_velocity position, EFFECT_GRAVITON, GRAVITON_TYPE_GIBLET, True, iIMDBaseShape_displayModel(headImd), col, gameTime - deltaGameTime + 1, velocity
    addEffect_velocity position, EFFECT_GRAVITON, GRAVITON_TYPE_GIBLET, True, iIMDBaseShape_displayModel(legsImd), col, gameTime - deltaGameTime + 1, velocity
    addEffect_velocity position, EFFECT_GRAVITON, GRAVITON_TYPE_GIBLET, True, iIMDBaseShape_displayModel(armImd), col, gameTime - deltaGameTime + 1, velocity
    addEffect_velocity position, EFFECT_GRAVITON, GRAVITON_TYPE_GIBLET, True, iIMDBaseShape_displayModel(bodyImd), col, gameTime - deltaGameTime + 1, velocity
End Sub


Public Function rescaleButtonObject(ByVal radius As Long, ByVal baseScale As Long, ByVal baseRadius As Long) As Long
    Dim newScale As Long
    newScale = 100 * baseRadius
    newScale = newScale \ radius
    If baseScale > 0 Then
        newScale = newScale + baseScale
        newScale = newScale \ 2
    End If
    rescaleButtonObject = newScale
End Function
