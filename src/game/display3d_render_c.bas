Attribute VB_Name = "display3d_render_c"
' src/display3d.cpp (part: lines 2499-3167, the objects' rendering: renderFeature,
' renderProximityMsg, getBlueprintColour, renderStructureTurrets,
' renderStructure, renderDeliveryPoint, renderWallSection). display3d.cpp's
' globals (playerPos, currentGameFrame, BLUEPRINT_OPACITY, ...) and its other
' static helpers (setScreenDispWithPerspective, structureBrightness,
' selectionBrightness, calcFlagPosScreenCoords, drawShape) are display3d_c's.
' An iIMDShape * / iIMDBaseShape * is WzGfx's handle, read through GfxApi's
' accessors; a mat4 goes to WzGfx by address (VarPtr). A PROXIMITY_DISPLAY *,
' MESSAGE *, VIEWDATA * and FLAG_POSITION * are message_c's and objmem_c's
' handles. A PIELIGHT is a Long (r in the low byte).
Option Explicit

' lib/ivis_opengl/pietypes.h's render style flags, imd.h's iV_IMD_NOSTRETCH
Private Const pie_ECM As Long = &H1
Private Const pie_TRANSLUCENT As Long = &H2
Private Const pie_ADDITIVE As Long = &H4
Private Const pie_FORCE_FOG As Long = &H8
Private Const pie_HEIGHT_SCALED As Long = &H10
Private Const pie_SHADOW As Long = &H80
Private Const pie_FORCELIGHT As Long = &H800
Private Const pie_RAISE_SCALE As Long = 256
Private Const iV_IMD_NOSTRETCH As Long = &H1000

' (display3d.h's BASE_MUZZLE_FLASH_DURATION: display3d_c's)

' (VB) glm::vec3(connectors[i].xzy()) of a display model; (0, 0, 0) past the end of the list, where C++
' reads past the vector (renderStructureTurrets indexes strImd->connectors[i] by weapon slot)
Private Function connectorXZY(ByVal shape As Long, ByVal i As Long) As Vector3f
    Dim x As Long, y As Long, z As Long
    If i < iIMDShape_connectors_size(shape) Then iIMDShape_connector shape, i, x, y, z
    connectorXZY = glm_vec3(CSng(x), CSng(z), CSng(y))
End Function

Private Sub connectorAt(ByVal shape As Long, ByVal i As Long, x As Long, y As Long, z As Long)
    x = 0
    y = 0
    z = 0
    If i < iIMDShape_connectors_size(shape) Then iIMDShape_connector shape, i, x, y, z
End Sub

' glm::translate(glm::vec3(dv)) of a Vector3i
Private Function translateI(dv As Vector3i) As mat4
    translateI = glm_translate(glm_vec3(CSng(dv.x), CSng(dv.y), CSng(dv.z)))
End Function

' PIELIGHT brightness.byte.r / g / b /= 2 (alpha kept)
Private Function halveRGB(ByVal c As Long) As Long
    halveRGB = ((c And &HFF&) \ 2) Or (((c And &HFF00&) \ &H100& \ 2) * &H100&) Or (((c And &HFF0000) \ &H10000 \ 2) * &H10000) Or (c And &HFF000000)
End Function

'/ Draw a feature (tree/rock/etc.)
Public Sub renderFeature(ByVal psFeature As Long, viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim brightness As Long, bForceDraw As Boolean, pieFlags As Long, dv As Vector3i, rotationV As Vector3i
    Dim modelMatrix As mat4, imd As Long, strImd As Long, stretchDepth As Single, pvm As mat4
    Dim psStats As Long, subType As FEATURE_TYPE

    brightness = pal_SetBrightness(200)
    psStats = aFeature(psFeature).psStats
    bForceDraw = (getRevealStatus() And asFeatureStats(psStats).visibleAtStart)
    pieFlags = 0

    If BASE_OBJECT_visibleForLocalDisplay(psFeature) = 0 And Not bForceDraw Then
        Exit Sub
    End If

    ' Mark it as having been drawn
    aFeature(psFeature).sDisplay.frameNumber = currentGameFrame

    ' Daft hack to get around the oil derrick issue
    If Not TileHasFeature(mapTile(map_coord(aFeature(psFeature).pos.x), map_coord(aFeature(psFeature).pos.y))) Then
        Exit Sub
    End If

    dv = Vector3i_( _
             aFeature(psFeature).pos.x, _
             aFeature(psFeature).pos.z, _
             -(aFeature(psFeature).pos.y))  ' features sits at the height of the tile it's centre is on

    ' Get all the pitch,roll,yaw info
    rotationV.y = -aFeature(psFeature).rot.direction
    rotationV.x = aFeature(psFeature).rot.pitch
    rotationV.z = aFeature(psFeature).rot.roll

    modelMatrix = glm_mat4_mul(glm_mat4_mul(glm_mat4_mul(translateI(dv), _
        glm_rotate(UNDEG(rotationV.y), glm_vec3(0!, 1!, 0!))), _
        glm_rotate(UNDEG(rotationV.x), glm_vec3(1!, 0!, 0!))), _
        glm_rotate(UNDEG(rotationV.z), glm_vec3(0!, 0!, 1!)))

    subType = asFeatureStats(psStats).subType
    If subType = FEAT_SKYSCRAPER Then
        modelMatrix = glm_mat4_mul(modelMatrix, objectShimmy(psFeature))
    End If

    If Not getRevealStatus() Then
        brightness = pal_SetBrightness(avGetObjLightLevel(psFeature, brightness And &HFF&))
    End If
    If Not hasSensorOnTile(mapTile(map_coord(aFeature(psFeature).pos.x), map_coord(aFeature(psFeature).pos.y)), selectedPlayer) Then
        brightness = halveRGB(brightness)
    End If

    If subType = FEAT_BUILDING _
       Or subType = FEAT_SKYSCRAPER _
       Or subType = FEAT_GEN_ARTE _
       Or subType = FEAT_BOULDER _
       Or subType = FEAT_VEHICLE _
       Or subType = FEAT_TREE _
       Or subType = FEAT_OIL_DRUM Then
        ' these cast a shadow
        pieFlags = pie_SHADOW
    End If

    imd = aFeature(psFeature).sDisplay.imd
    If imd <> 0 Then
        strImd = iIMDBaseShape_displayModel(imd)

        stretchDepth = 0!
        If (iIMDShape_flags(strImd) And iV_IMD_NOSTRETCH) = 0 Then
            If subType = FEAT_TREE Then ' for now, only do this for trees
                stretchDepth = aFeature(psFeature).pos.z - aFeature(psFeature).foundationDepth
            End If
        End If

        ' Translate the feature  - N.B. We can also do rotations here should we require
        ' buildings to face different ways - Don't know if this is necessary - should be IMO
        pie_Draw3DShape strImd, 0, 0, brightness, pieFlags, 0, VarPtr(modelMatrix), VarPtr(viewMatrix), stretchDepth
    End If

    pvm = glm_mat4_mul(perspectiveViewMatrix, modelMatrix)
    setScreenDispWithPerspective aFeature(psFeature).sDisplay, pvm
End Sub

Public Sub renderProximityMsg(ByVal psProxDisp As Long, viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim msgX As Long, msgY As Long, dv As Vector3i, pViewProximity As Long, x As Long, y As Long, r As Long
    Dim proxImd As Long, psMessage As Long, pViewData As Long, psObj As Long, objPos As Vector3i
    Dim modelMatrix As mat4, pvm As mat4
    msgX = 0
    msgY = 0
    dv = Vector3i_(0, 0, 0)
    pViewProximity = 0
    proxImd = 0

    'store the frame number for when deciding what has been clicked on
    aProxDisp(psProxDisp).frameNumber = currentGameFrame

    ' Get it's x and y coordinates so we don't have to deref. struct later
    psMessage = aProxDisp(psProxDisp).psMessage
    If aProxDisp(psProxDisp).typ = POS_PROXDATA Then
        pViewData = aMessage(psMessage).pViewData
        If pViewData = 0 Then
            Exit Sub ' if no data - ignore message
        End If
        If aViewData(pViewData).hasProxData Then pViewProximity = pViewData   ' (VIEW_PROXIMITY *)pViewData->pData
        If pViewProximity <> 0 Then
            msgX = aViewData(pViewProximity).prox_x
            msgY = aViewData(pViewProximity).prox_y
            ' message sits at the height specified at input
            dv.y = aViewData(pViewProximity).prox_z + 64

            ' in case of a beacon message put above objects
            If aViewData(pViewData).typ = VIEW_BEACON Then
                If TileIsOccupied(mapTile(FromU32(Int(U32(msgX) / TILE_UNITS)), FromU32(Int(U32(msgY) / TILE_UNITS)))) Then
                    dv.y = aViewData(pViewProximity).prox_z + 150
                End If
            End If

        End If
    ElseIf aProxDisp(psProxDisp).typ = POS_PROXOBJ Then
        psObj = aMessage(psMessage).psObj
        objPos = SIMPLE_OBJECT_pos(psObj)
        msgX = objPos.x
        msgY = objPos.y
        ' message sits at the height specified at input
        dv.y = objPos.z + 64
    Else
        ASSERT_FAILURE "!""unknown proximity display message type""", "src/display3d.cpp:2632", "renderProximityMsg", "Buggered proximity message type"
        Exit Sub
    End If

    dv.x = msgX
    dv.z = -msgY

    ' Translate the message
    modelMatrix = translateI(dv)

    'get the appropriate IMD
    If pViewProximity <> 0 Then
        Select Case aViewData(pViewProximity).prox_proxType
        Case PROX_ENEMY
            proxImd = getDisplayImdFromIndex(MI_BLIP_ENEMY)
        Case PROX_RESOURCE
            proxImd = getDisplayImdFromIndex(MI_BLIP_RESOURCE)
        Case PROX_ARTEFACT
            proxImd = getDisplayImdFromIndex(MI_BLIP_ARTEFACT)
        Case Else
            ASSERT_FAILURE "!""unknown proximity display message type""", "src/display3d.cpp:2664", "renderProximityMsg", "Buggered proximity message type"
        End Select
    Else
        'object Proximity displays are for oil resources and artefacts
        If Not (SIMPLE_OBJECT_type(aMessage(psMessage).psObj) = OBJ_FEATURE) Then ASSERT_FAILURE "psProxDisp->psMessage->psObj->type == OBJ_FEATURE", "src/display3d.cpp:2671", "renderProximityMsg", "Invalid object type for proximity display": Exit Sub

        If asFeatureStats(aFeature(castFeature(aMessage(psMessage).psObj)).psStats).subType = FEAT_OIL_RESOURCE Then
            'resource
            proxImd = getDisplayImdFromIndex(MI_BLIP_RESOURCE)
        Else
            'artefact
            proxImd = getDisplayImdFromIndex(MI_BLIP_ARTEFACT)
        End If
    End If

    modelMatrix = glm_mat4_mul(modelMatrix, glm_mat4_mul(glm_rotate(UNDEG(-playerPos.r.y), glm_vec3(0!, 1!, 0!)), _
        glm_rotate(UNDEG(-playerPos.r.x), glm_vec3(1!, 0!, 0!))))

    If proxImd <> 0 Then
        pie_Draw3DShape proxImd, getModularScaledGraphicsTime(iIMDShape_animInterval(proxImd), iIMDShape_numFrames(proxImd)), 0, psPalette(WZCOL_WHITE), pie_ADDITIVE, 192, VarPtr(modelMatrix), VarPtr(viewMatrix)
    End If
    'get the screen coords for determining when clicked on
    pvm = glm_mat4_mul(perspectiveViewMatrix, modelMatrix)
    calcFlagPosScreenCoords x, y, r, pvm
    aProxDisp(psProxDisp).screenX = x
    aProxDisp(psProxDisp).screenY = y
    aProxDisp(psProxDisp).screenR = r
End Sub

' (VB) Public: display3d_c's structureBrightness uses it
Public Function getBlueprintColour(ByVal state As STRUCT_STATES) As Long
    Select Case state
    Case SS_BLUEPRINT_VALID
        getBlueprintColour = psPalette(WZCOL_LGREEN)
    Case SS_BLUEPRINT_INVALID
        getBlueprintColour = psPalette(WZCOL_LRED)
    Case SS_BLUEPRINT_PLANNED
        getBlueprintColour = psPalette(WZCOL_BLUEPRINT_PLANNED)
    Case SS_BLUEPRINT_PLANNED_BY_ALLY
        getBlueprintColour = psPalette(WZCOL_BLUEPRINT_PLANNED_BY_ALLY)
    Case Else
        debug_ 2712, LOG_ERROR, "getBlueprintColour", "this is not a blueprint"
        getBlueprintColour = psPalette(WZCOL_WHITE)
    End Select
End Function

Private Sub renderStructureTurrets(ByVal psStructure As Long, ByVal strImd As Long, ByVal buildingBrightness As Long, ByVal pieFlag As Long, ByVal pieFlagData As Long, ByVal ecmFlag As Long, _
    modelMatrix As mat4, viewMatrix As mat4)
    Dim mountImd(0 To MAX_WEAPONS - 1) As Long, weaponImd(0 To MAX_WEAPONS - 1) As Long, flashImd(0 To MAX_WEAPONS - 1) As Long
    Dim colour As Long, i As Long, n As Long, nWeaponStat As Long, noRecoil As Boolean, pStats As Long
    Dim rot As Rotation, matrix As mat4, heightAboveTerrain As Single, recoilValue As Long
    Dim pMountDisplayIMD As Long, animFrame As Long, pWeaponDisplayIMD As Long, mm As mat4
    Dim psRepairObj As Long, psDroid As Long, xdiff As Long, ydiff As Long, pRepImd As Long
    Dim cx As Long, cy As Long, cz As Long, pFlashDisplayIMD As Long, frame As Long, mxx As Long, mxy As Long, mxz As Long
    Dim cIdx As Long, lImd As Long, structDisplay As Long, nConnectors As Long

    colour = getPlayerColour(aStruct(psStructure).player)
    pStats = aStruct(psStructure).pStructureType

    'get an imd to draw on the connector priority is weapon, ECM, sensor
    'check for weapon
    If aStruct(psStructure).numWeaps > 1 Then n = aStruct(psStructure).numWeaps Else n = 1
    For i = 0 To n - 1
        If aStruct(psStructure).asWeaps(i).nStat > 0 Then
            nWeaponStat = aStruct(psStructure).asWeaps(i).nStat

            weaponImd(i) = asWeaponStats(nWeaponStat).pIMD
            mountImd(i) = asWeaponStats(nWeaponStat).pMountGraphic
            flashImd(i) = asWeaponStats(nWeaponStat).pMuzzleGraphic
        End If
    Next

    ' check for ECM
    If weaponImd(0) = 0 And asStructureStats(pStats).pECM >= 0 Then
        weaponImd(0) = asECMStats(asStructureStats(pStats).pECM).pIMD
        mountImd(0) = asECMStats(asStructureStats(pStats).pECM).pMountGraphic
        flashImd(0) = 0
    End If
    ' check for sensor (or repair center)
    noRecoil = False
    If weaponImd(0) = 0 And asStructureStats(pStats).pSensor >= 0 Then
        weaponImd(0) = asSensorStats(asStructureStats(pStats).pSensor).pIMD
        ' No recoil for sensors
        noRecoil = True
        mountImd(0) = asSensorStats(asStructureStats(pStats).pSensor).pMountGraphic
        flashImd(0) = 0
    End If

    ' flags for drawing weapons
    If STRUCTURE_isBlueprint(psStructure) Then
        pieFlag = pie_TRANSLUCENT
        pieFlagData = BLUEPRINT_OPACITY
    Else
        pieFlag = pie_SHADOW Or ecmFlag
        pieFlagData = 0
    End If

    ' draw Weapon / ECM / Sensor for structure
    i = 0
    Do While i < aStruct(psStructure).numWeaps Or i = 0
        rot = structureGetInterpolatedWeaponRotation(psStructure, i, graphicsTime)

        If weaponImd(i) <> 0 Then
            connectorAt strImd, i, cx, cy, cz
            matrix = glm_mat4_mul(glm_translate(glm_vec3(CSng(cx), CSng(cz), CSng(cy))), glm_rotate(UNDEG(-rot.direction), glm_vec3(0!, 1!, 0!)))
            heightAboveTerrain = cz
            If noRecoil Then recoilValue = 0 Else recoilValue = getRecoil(aStruct(psStructure).asWeaps(i))
            If mountImd(i) <> 0 Then
                pMountDisplayIMD = iIMDBaseShape_displayModel(mountImd(i))
                matrix = glm_mat4_mul(matrix, glm_translate(glm_vec3(0!, 0!, recoilValue / 3!)))
                animFrame = 0
                If iIMDShape_numFrames(pMountDisplayIMD) > 0 Then ' Calculate an animation frame
                    animFrame = getModularScaledGraphicsTime(iIMDShape_animInterval(pMountDisplayIMD), iIMDShape_numFrames(pMountDisplayIMD))
                End If
                mm = glm_mat4_mul(modelMatrix, matrix)
                pie_Draw3DShape pMountDisplayIMD, animFrame, colour, buildingBrightness, pieFlag, pieFlagData, VarPtr(mm), VarPtr(viewMatrix), -heightAboveTerrain
                If iIMDShape_connectors_size(pMountDisplayIMD) <> 0 Then
                    matrix = glm_mat4_mul(matrix, glm_translate(connectorXZY(pMountDisplayIMD, 0)))
                End If
            End If
            matrix = glm_mat4_mul(matrix, glm_rotate(UNDEG(rot.pitch), glm_vec3(1!, 0!, 0!)))
            matrix = glm_mat4_mul(matrix, glm_translate(glm_vec3(0!, 0!, CSng(recoilValue))))

            pWeaponDisplayIMD = iIMDBaseShape_displayModel(weaponImd(i))
            mm = glm_mat4_mul(modelMatrix, matrix)
            pie_Draw3DShape pWeaponDisplayIMD, 0, colour, buildingBrightness, pieFlag, pieFlagData, VarPtr(mm), VarPtr(viewMatrix), -heightAboveTerrain
            If aStruct(psStructure).status = SS_BUILT And BASE_OBJECT_visibleForLocalDisplay(psStructure) > (UBYTE_MAX \ 2) Then
                If asStructureStats(pStats).typ = REF_REPAIR_FACILITY Then
                    ' REPAIR_FACILITY *psRepairFac = &psStructure->pFunctionality->repairFacility;
                    psRepairObj = aStruct(psStructure).functionality.repairFacility.psObj
                    ' draw repair flash if the Repair Facility has a target which it has started work on
                    If iIMDShape_connectors_size(pWeaponDisplayIMD) <> 0 And psRepairObj <> 0 Then
                        If SIMPLE_OBJECT_type(psRepairObj) = OBJ_DROID Then
                            psDroid = psRepairObj
                            xdiff = aDroid(psDroid).pos.x - aStruct(psStructure).pos.x
                            ydiff = aDroid(psDroid).pos.y - aStruct(psStructure).pos.y
                            If xdiff * xdiff + ydiff * ydiff <= (TILE_UNITS * 5 \ 2) * (TILE_UNITS * 5 \ 2) Then
                                pRepImd = getDisplayImdFromIndex(MI_FLAME)

                                iIMDShape_connector pWeaponDisplayIMD, 0, cx, cy, cz
                                matrix = glm_mat4_mul(matrix, glm_mat4_mul(glm_mat4_mul(glm_mat4_mul( _
                                    glm_translate(glm_vec3(CSng(cx), CSng(cz - 12), CSng(cy))), _
                                    glm_rotate(UNDEG(rot.direction), glm_vec3(0!, 1!, 0!))), _
                                    glm_rotate(UNDEG(-playerPos.r.y), glm_vec3(0!, 1!, 0!))), _
                                    glm_rotate(UNDEG(-playerPos.r.x), glm_vec3(1!, 0!, 0!))))
                                mm = glm_mat4_mul(modelMatrix, matrix)
                                pie_Draw3DShape pRepImd, getModularScaledGraphicsTime(iIMDShape_animInterval(pRepImd), iIMDShape_numFrames(pRepImd)), colour, buildingBrightness, pie_ADDITIVE, 192, VarPtr(mm), VarPtr(viewMatrix), -heightAboveTerrain
                            End If
                        End If
                    End If
                Else ' we have a weapon so we draw a muzzle flash
                    If flashImd(i) <> 0 Then pFlashDisplayIMD = iIMDBaseShape_displayModel(flashImd(i)) Else pFlashDisplayIMD = 0
                    mm = glm_mat4_mul(modelMatrix, matrix)
                    drawMuzzleFlash aStruct(psStructure).asWeaps(i), pWeaponDisplayIMD, pFlashDisplayIMD, buildingBrightness, pieFlag, pieFlagData, mm, viewMatrix, heightAboveTerrain, colour
                End If
            End If
        ' no IMD, its a baba machine gun, bunker, etc.
        ElseIf aStruct(psStructure).asWeaps(i).nStat > 0 Then
            If aStruct(psStructure).status = SS_BUILT Then
                nWeaponStat = aStruct(psStructure).asWeaps(i).nStat

                ' get an imd to draw on the connector priority is weapon, ECM, sensor
                ' check for weapon
                flashImd(i) = asWeaponStats(nWeaponStat).pMuzzleGraphic
                ' draw Weapon/ECM/Sensor for structure
                If flashImd(i) <> 0 Then
                    rot = aStruct(psStructure).asWeaps(i).rot ' Snap rotation so the muzzle graphic appears where it should, see aiUpdateStructure().
                    pFlashDisplayIMD = iIMDBaseShape_displayModel(flashImd(i))
                    matrix = glm_mat4(1!)
                    ' horrendous hack
                    iIMDShape_max strImd, mxx, mxy, mxz
                    If mxy > 80 Then ' babatower
                        matrix = glm_mat4_mul(matrix, glm_mat4_mul(glm_mat4_mul(glm_translate(glm_vec3(0!, 80!, 0!)), glm_rotate(UNDEG(-rot.direction), glm_vec3(0!, 1!, 0!))), glm_translate(glm_vec3(0!, 0!, -20!))))
                    Else 'baba bunker
                        matrix = glm_mat4_mul(matrix, glm_mat4_mul(glm_mat4_mul(glm_translate(glm_vec3(0!, 10!, 0!)), glm_rotate(UNDEG(-rot.direction), glm_vec3(0!, 1!, 0!))), glm_translate(glm_vec3(0!, 0!, -40!))))
                    End If
                    matrix = glm_mat4_mul(matrix, glm_rotate(UNDEG(rot.pitch), glm_vec3(1!, 0!, 0!)))
                    ' draw the muzzle flash?
                    If BASE_OBJECT_visibleForLocalDisplay(psStructure) > UBYTE_MAX \ 2 Then
                        ' animate for the duration of the flash only
                        ' assume no clan colours for muzzle effects
                        If iIMDShape_numFrames(pFlashDisplayIMD) = 0 Or iIMDShape_animInterval(pFlashDisplayIMD) <= 0 Then
                            ' no anim so display one frame for a fixed time
                            If Not ULT(graphicsTime, aStruct(psStructure).asWeaps(i).lastFired) And ULT(graphicsTime, aStruct(psStructure).asWeaps(i).lastFired + BASE_MUZZLE_FLASH_DURATION) Then
                                mm = glm_mat4_mul(modelMatrix, matrix)
                                pie_Draw3DShape pFlashDisplayIMD, 0, colour, buildingBrightness, 0, 0, VarPtr(mm), VarPtr(viewMatrix) 'muzzle flash
                            End If
                        Else
                            frame = FromU32(Int(U32(graphicsTime - aStruct(psStructure).asWeaps(i).lastFired) / iIMDShape_animInterval(pFlashDisplayIMD)))
                            If frame < iIMDShape_numFrames(pFlashDisplayIMD) And frame >= 0 Then
                                mm = glm_mat4_mul(modelMatrix, matrix)
                                pie_Draw3DShape pFlashDisplayIMD, 0, colour, buildingBrightness, 0, 0, VarPtr(mm), VarPtr(viewMatrix) 'muzzle flash
                            End If
                        End If
                    End If
                End If
            End If
        ' if there is an unused connector, but not the first connector, add a light to it
        Else
            structDisplay = iIMDBaseShape_displayModel(aStruct(psStructure).sDisplay.imd)
            nConnectors = iIMDShape_connectors_size(structDisplay)
            If nConnectors > 1 Then
                Select Case asStructureStats(pStats).typ
                Case REF_FACTORY, REF_FACTORY_MODULE, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
                    ' don't do this
                Case Else
                    For cIdx = 1 To nConnectors - 1
                        lImd = getDisplayImdFromIndex(MI_LANDING)
                        mm = glm_mat4_mul(modelMatrix, glm_translate(connectorXZY(structDisplay, cIdx)))
                        pie_Draw3DShape lImd, getModularScaledGraphicsTime(iIMDShape_animInterval(lImd), iIMDShape_numFrames(lImd)), colour, buildingBrightness, 0, 0, _
                                        VarPtr(mm), VarPtr(viewMatrix)
                    Next
                End Select
            End If
        End If
        i = i + 1
    Loop
End Sub

'/ Draw the structures
Public Sub renderStructure(ByVal psStructure As Long, viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim colour As Long, pieFlagData As Long, ecmFlag As Long, pieFlag As Long, buildingBrightness As Long
    Dim dv As Vector3i, bHitByElectronic As Boolean, defensive As Boolean, faction_ As Long, strImd As Long
    Dim psTile As Long, modelMatrix As mat4, visibilityAmount As Long, frame As Long, m As mat4, imd As Long
    Dim pStats As Long, stretch As Single, pvm As mat4

    ecmFlag = 0
    pieFlag = 0
    dv = Vector3i_(aStruct(psStructure).pos.x, aStruct(psStructure).pos.z, -(aStruct(psStructure).pos.y))
    bHitByElectronic = False
    defensive = False
    faction_ = getPlayerFaction(aStruct(psStructure).player)
    strImd = getFactionDisplayIMD(faction_, iIMDBaseShape_displayModel(aStruct(psStructure).sDisplay.imd))
    psTile = worldTile(aStruct(psStructure).pos.x, aStruct(psStructure).pos.y)
    pStats = aStruct(psStructure).pStructureType

    modelMatrix = glm_mat4_mul(translateI(dv), glm_rotate(UNDEG(-aStruct(psStructure).rot.direction), glm_vec3(0!, 1!, 0!)))

    If asStructureStats(pStats).typ = REF_WALL Or asStructureStats(pStats).typ = REF_WALLCORNER _
       Or asStructureStats(pStats).typ = REF_GATE Then
        renderWallSection psStructure, viewMatrix, perspectiveViewMatrix
        Exit Sub
    End If
    ' If the structure is not truly visible, but we know there is something there, we will instead draw a blip
    visibilityAmount = BASE_OBJECT_visibleForLocalDisplay(psStructure)
    If visibilityAmount < UBYTE_MAX And visibilityAmount > 0 Then
        frame = FromU32(Int(U32(graphicsTime) / BLIP_ANIM_DURATION) + (U32(aStruct(psStructure).id) - Int(U32(aStruct(psStructure).id) / 8192#) * 8192#))  ' de-sync the blip effect, but don't overflow the int
        m = translateI(dv)
        pie_Draw3DShape getFactionDisplayIMD(faction_, getDisplayImdFromIndex(MI_BLIP)), frame, 0, psPalette(WZCOL_WHITE), pie_ADDITIVE, visibilityAmount \ 2, _
            VarPtr(m), VarPtr(viewMatrix)
        Exit Sub
    ElseIf visibilityAmount = 0 Then
        Exit Sub
    ElseIf asStructureStats(pStats).typ = REF_DEFENSE Then
        defensive = True
    End If

    If (psMapTiles(psTile).jammerBits And alliancebits(aStruct(psStructure).player)) <> 0 Then
        ecmFlag = pie_ECM
    End If

    colour = getPlayerColour(aStruct(psStructure).player)

    ' -------------------------------------------------------------------------------

    ' Mark it as having been drawn
    aStruct(psStructure).sDisplay.frameNumber = currentGameFrame

    If Not defensive Then
        If ULT(aStruct(psStructure).timeLastHit - graphicsTime, ELEC_DAMAGE_DURATION) _
           And aStruct(psStructure).lastHitWeapon = WSC_ELECTRONIC Then
            bHitByElectronic = True
        End If
    End If

    buildingBrightness = structureBrightness(psStructure)

    If Not defensive Then
        ' Draw the building's base first
        If asStructureStats(pStats).pBaseIMD <> 0 Then
            If STRUCTURE_isBlueprint(psStructure) Then
                pieFlagData = BLUEPRINT_OPACITY
            Else
                pieFlag = pie_FORCE_FOG Or ecmFlag
                pieFlagData = 255
            End If
            pie_Draw3DShape getFactionDisplayIMD(faction_, iIMDBaseShape_displayModel(asStructureStats(pStats).pBaseIMD)), 0, colour, buildingBrightness, pieFlag Or pie_TRANSLUCENT Or pie_FORCELIGHT, pieFlagData, _
                VarPtr(modelMatrix), VarPtr(viewMatrix)
        End If

        ' override
        If bHitByElectronic Then
            buildingBrightness = pal_SetBrightness(150)
        End If
    End If

    If bHitByElectronic Then
        modelMatrix = glm_mat4_mul(modelMatrix, objectShimmy(psStructure))
    End If

    'first check if partially built - ANOTHER HACK!
    If aStruct(psStructure).status = SS_BEING_BUILT Then
        If aStruct(psStructure).prebuiltImd <> 0 Then
            ' strImd is a module, so render the already-built part at full height.
            imd = aStruct(psStructure).prebuiltImd
            If imd <> 0 Then
                pie_Draw3DShape getFactionDisplayIMD(faction_, iIMDBaseShape_displayModel(imd)), 0, colour, buildingBrightness, pie_SHADOW, 0, _
                    VarPtr(modelMatrix), VarPtr(viewMatrix)
            End If
        End If
        If strImd <> 0 Then
            pie_Draw3DShape strImd, 0, colour, buildingBrightness, pie_HEIGHT_SCALED Or pie_SHADOW, CLng(Fix(structHeightScale(psStructure) * pie_RAISE_SCALE)), VarPtr(modelMatrix), VarPtr(viewMatrix)
        End If
        pvm = glm_mat4_mul(perspectiveViewMatrix, modelMatrix)
        setScreenDispWithPerspective aStruct(psStructure).sDisplay, pvm
        Exit Sub
    End If

    If STRUCTURE_isBlueprint(psStructure) Then
        pieFlag = pie_TRANSLUCENT
        pieFlagData = BLUEPRINT_OPACITY
    Else
        ' structures can be rotated, so use a dynamic shadow for them
        pieFlag = pie_SHADOW Or ecmFlag
        pieFlagData = 0
    End If

    ' check for animation model replacement - if none found, use animation in existing IMD
    If iIMDShape_objanimpie(strImd, aStruct(psStructure).animationEvent) <> 0 Then
        strImd = iIMDBaseShape_displayModel(iIMDShape_objanimpie(strImd, aStruct(psStructure).animationEvent))
    End If

    Do While strImd <> 0
        stretch = 0!
        If defensive And Not STRUCTURE_isBlueprint(psStructure) Then
            If (iIMDShape_flags(strImd) And iV_IMD_NOSTRETCH) = 0 Then
                stretch = aStruct(psStructure).pos.z - aStruct(psStructure).foundationDepth
            End If
        End If
        drawShape strImd, aStruct(psStructure).timeAnimationStarted, colour, buildingBrightness, pieFlag, pieFlagData, modelMatrix, viewMatrix, stretch
        If iIMDShape_connectors_size(strImd) > 0 Then
            renderStructureTurrets psStructure, strImd, buildingBrightness, pieFlag, pieFlagData, ecmFlag, modelMatrix, viewMatrix
        End If
        strImd = iIMDShape_next(strImd)
    Loop
    pvm = glm_mat4_mul(perspectiveViewMatrix, modelMatrix)
    setScreenDispWithPerspective aStruct(psStructure).sDisplay, pvm
End Sub

'/ draw the delivery points (psPosition: a FLAG_POSITION *)
Public Sub renderDeliveryPoint(ByVal psPosition As Long, ByVal blueprint_ As Boolean, viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim dv As Vector3i, x As Long, y As Long, r As Long, pieFlag As Long, pieFlagData As Long, colour As Long
    Dim modelMatrix As mat4, structure_ As Long, psIMD As Long, pvm As mat4

    'store the frame number for when deciding what has been clicked on
    aFlagPos(psPosition).frameNumber = currentGameFrame

    dv.x = aFlagPos(psPosition).coords.x
    dv.z = -(aFlagPos(psPosition).coords.y)
    dv.y = aFlagPos(psPosition).coords.z

    'quick check for invalid data
    If Not (psPosition <> 0 And aFlagPos(psPosition).factoryType < NUM_FLAG_TYPES And aFlagPos(psPosition).factoryInc < MAX_FACTORY_FLAG_IMDS) Then
        ASSERT_FAILURE "psPosition && psPosition->factoryType < NUM_FLAG_TYPES && psPosition->factoryInc < MAX_FACTORY_FLAG_IMDS", "src/display3d.cpp:3068", "renderDeliveryPoint", "Invalid assembly point"
        Exit Sub
    End If

    modelMatrix = glm_mat4_mul(glm_mat4_mul(translateI(dv), glm_scale(glm_vec3(0.5!, 0.5!, 0.5!))), glm_rotate(-UNDEG(playerPos.r.y), glm_vec3(0!, 1!, 0!)))

    pieFlag = pie_TRANSLUCENT
    pieFlagData = BLUEPRINT_OPACITY

    If blueprint_ Then
        If deliveryReposValid() Then colour = psPalette(WZCOL_LGREEN) Else colour = psPalette(WZCOL_LRED)
    Else
        pieFlag = pieFlag Or pie_FORCE_FOG
        colour = psPalette(WZCOL_WHITE)
        structure_ = findDeliveryFactory(psPosition)
        If structure_ <> 0 Then
            If aStruct(structure_).selected <> 0 Then
                colour = selectionBrightness()
            End If
        End If
    End If
    psIMD = pAssemblyPointIMDs(aFlagPos(psPosition).factoryType, aFlagPos(psPosition).factoryInc)
    If psIMD <> 0 Then
        pie_Draw3DShape iIMDBaseShape_displayModel(psIMD), 0, 0, colour, pieFlag, pieFlagData, VarPtr(modelMatrix), VarPtr(viewMatrix)
    End If

    'get the screen coords for the DP
    pvm = glm_mat4_mul(perspectiveViewMatrix, modelMatrix)
    calcFlagPosScreenCoords x, y, r, pvm
    aFlagPos(psPosition).screenX = x
    aFlagPos(psPosition).screenY = y
    aFlagPos(psPosition).screenR = r
End Sub

Private Function renderWallSection(ByVal psStructure As Long, viewMatrix As mat4, perspectiveViewMatrix As mat4) As Boolean
    Dim ecmFlag As Long, brightness As Long, dv As Vector3i, pieFlag As Long, pieFlagData As Long
    Dim psTile As Long, faction_ As Long, stretch As Single, modelMatrix As mat4, imd As Long, pvm As mat4
    ecmFlag = 0
    psTile = worldTile(aStruct(psStructure).pos.x, aStruct(psStructure).pos.y)
    faction_ = getPlayerFaction(aStruct(psStructure).player)

    If BASE_OBJECT_visibleForLocalDisplay(psStructure) = 0 Then
        renderWallSection = False
        Exit Function
    End If

    If (psMapTiles(psTile).jammerBits And alliancebits(aStruct(psStructure).player)) <> 0 Then
        ecmFlag = pie_ECM
    End If

    aStruct(psStructure).sDisplay.frameNumber = currentGameFrame

    brightness = structureBrightness(psStructure)
    stretch = aStruct(psStructure).pos.z - aStruct(psStructure).foundationDepth

    ' Establish where it is in the world
    dv.x = aStruct(psStructure).pos.x
    dv.z = -(aStruct(psStructure).pos.y)
    dv.y = aStruct(psStructure).pos.z

    dv.y = dv.y - gateCurrentOpenHeight(psStructure, graphicsTime, 1)  ' Make gate stick out by 1 unit, so that the tops of + gates can safely have heights differing by 1 unit.

    modelMatrix = glm_mat4_mul(translateI(dv), glm_rotate(UNDEG(-aStruct(psStructure).rot.direction), glm_vec3(0!, 1!, 0!)))

    ' Actually render it
    If aStruct(psStructure).status = SS_BEING_BUILT Then
        If aStruct(psStructure).sDisplay.imd <> 0 Then
            pie_Draw3DShape getFactionDisplayIMD(faction_, iIMDBaseShape_displayModel(aStruct(psStructure).sDisplay.imd)), 0, getPlayerColour(aStruct(psStructure).player), _
                            brightness, pie_HEIGHT_SCALED Or pie_SHADOW Or ecmFlag, CLng(Fix(structHeightScale(psStructure) * pie_RAISE_SCALE)), VarPtr(modelMatrix), VarPtr(viewMatrix), stretch
        End If
    Else
        If STRUCTURE_isBlueprint(psStructure) Then
            pieFlag = pie_TRANSLUCENT
            pieFlagData = BLUEPRINT_OPACITY
        Else
            ' Use a dynamic shadow
            pieFlag = pie_SHADOW
            pieFlagData = 0
        End If
        imd = aStruct(psStructure).sDisplay.imd
        If imd <> 0 Then
            pie_Draw3DShape getFactionDisplayIMD(faction_, iIMDBaseShape_displayModel(imd)), 0, getPlayerColour(aStruct(psStructure).player), brightness, pieFlag Or ecmFlag, pieFlagData, VarPtr(modelMatrix), VarPtr(viewMatrix), stretch
        End If
    End If
    pvm = glm_mat4_mul(perspectiveViewMatrix, modelMatrix)
    setScreenDispWithPerspective aStruct(psStructure).sDisplay, pvm
    renderWallSection = True
End Function
