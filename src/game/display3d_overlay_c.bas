Attribute VB_Name = "display3d_overlay_c"
' src/display3d.cpp (part: lines 400-583, 3168-3924, 4103-4637): the 2D
' overlays drawn over the 3D view - BatchedObjectStatusRenderer (the
' health, power and reload bars, batched in WzGfx's
' BatchedMultiRectRenderer), the drag box, droid and structure
' selections, group, rank and commander numbers, the sensor and
' destination targets, the power generator and rearm pad effects, the
' sensor and weapon ranges, the construction lines. display3d.cpp's
' file-level globals (barMode, currentGameFrame, tuiTargetOrigin,
' bSensorTargetting, the range display's, underwaterTile, ...) are
' display3d_c's. The class BatchedObjectStatusRenderer has one instance
' (batchedObjectStatusRenderer): its methods are
' batchedObjectStatusRenderer_<method>, its members this module's.
' The static functions draw3DScene calls (display3d_c) are Public.
Option Explicit

' X/Y offset to display the group number at
Private Const GN_X_OFFSET As Long = 8
Private Const GN_Y_OFFSET As Long = 3

Private Enum GROUPNUMBER_TYPE
    GN_NORMAL
    GN_DAMAGED
    GN_FACTORY
End Enum

' X offset to display the group number at
Private Const CMND_STAR_X_OFFSET As Long = 6
Private Const CMND_GN_Y_OFFSET As Long = 8

' BatchedObjectStatusRenderer's members
Private droidsToDraw() As Long          ' std::vector<DROID *>
Private droidsToDraw_size As Long
Private structsToDraw() As Long         ' std::vector<STRUCTURE *>
Private structsToDraw_size As Long
Private renderStructureTargetingGfx As Boolean ' = false
Private batchedMultiRectRenderer As BatchedMultiRectRenderer

' (VB) the vectors' first storage
Public Sub display3d_overlay_c_init()
    ReDim droidsToDraw(0 To 63)
    ReDim structsToDraw(0 To 63)
    droidsToDraw_size = 0
    structsToDraw_size = 0
    renderStructureTargetingGfx = False
    Set batchedMultiRectRenderer = New BatchedMultiRectRenderer
End Sub

' (VB) PERCENT(a, b): (a) * 100 / (b), as the UDWORD arithmetic of its uses
Private Function PERCENT_(ByVal a As Long, ByVal b As Long) As Long
    PERCENT_ = FromU32(Fix((U32(a) * 100) / U32(b)))
End Function

' (VB) BatchedObjectStatusRenderer::addDroidSelectionsToRender
Private Sub batchedObjectStatusRenderer_addDroidSelectionsToRender(ByVal player As Long, ByVal psClickedOn As Long, ByVal bMouseOverOwnDroid As Boolean)
    Dim bDrawAll As Boolean, it As Long, psDroid As Long, i As Long, draw As Boolean
    bDrawAll = barMode = BAR_DROIDS Or barMode = BAR_DROIDS_AND_STRUCTURES

    ' first, for all droids, queue the various rects
    it = apsDroidLists(player).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).sDisplay.frameNumber = currentGameFrame Then
            ' If it's selected and on screen or it's the one the mouse is over
            draw = bDrawAll
            If Not draw Then draw = eitherSelected(psDroid)
            If Not draw Then draw = (bMouseOverOwnDroid And psDroid = psClickedOn)
            If Not draw Then draw = droidUnderRepair(psDroid)
            If draw Then
                queueDroidPowerBarsRects psDroid, aDroid(psDroid).selected, batchedMultiRectRenderer, 0

                For i = 0 To aDroid(psDroid).numWeaps - 1
                    queueWeaponReloadBarRects psDroid, i, batchedMultiRectRenderer, 1
                Next

                If droidsToDraw_size > UBound(droidsToDraw) Then ReDim Preserve droidsToDraw(0 To 2 * droidsToDraw_size)
                droidsToDraw(droidsToDraw_size) = psDroid
                droidsToDraw_size = droidsToDraw_size + 1
            End If
        End If
        it = lnNext(it)
    Loop
End Sub

Private Sub batchedObjectStatusRenderer_addEnemyDroidHealthToRender(ByVal psDroid As Long)
    queueDroidEnemyHealthBarsRects psDroid, batchedMultiRectRenderer, 4
End Sub

Private Sub batchedObjectStatusRenderer_addEnemyStructureHealthToRender(ByVal psStruct As Long)
    queueStructureHealth psStruct, batchedMultiRectRenderer, 4
    If aStruct(psStruct).status = SS_BEING_BUILT Then
        queueStructureBuildProgress psStruct, batchedMultiRectRenderer, 4
    End If
End Sub

Private Sub batchedObjectStatusRenderer_addStructureSelectionsToRender(ByVal player As Long, ByVal psClickedOn As Long, ByVal bMouseOverOwnStructure As Boolean)
    Dim it As Long, psStruct As Long, i As Long, draw As Boolean, typ As Long
    ' Go thru' all the buildings
    it = apsStructLists(player).head
    Do While it <> 0
        psStruct = lnObj(it)
        If aStruct(psStruct).sDisplay.frameNumber = currentGameFrame Then
            ' If it's selected
            typ = asStructureStats(aStruct(psStruct).pStructureType).typ
            draw = aStruct(psStruct).selected Or _
                   (barMode = BAR_DROIDS_AND_STRUCTURES And typ <> REF_WALL And typ <> REF_WALLCORNER) Or _
                   (bMouseOverOwnStructure And psStruct = psClickedOn)
            If draw Then
                queueStructureHealth psStruct, batchedMultiRectRenderer, 2

                For i = 0 To aStruct(psStruct).numWeaps - 1
                    queueWeaponReloadBarRects psStruct, i, batchedMultiRectRenderer, 3
                    ' drawStructureTargetOriginIcon(psStruct, i); // in the future, use instanced rendering here? (for now, add to structsToDraw below and handle later)
                Next

                If structsToDraw_size > UBound(structsToDraw) Then ReDim Preserve structsToDraw(0 To 2 * structsToDraw_size)
                structsToDraw(structsToDraw_size) = psStruct
                structsToDraw_size = structsToDraw_size + 1
            End If

            If aStruct(psStruct).status = SS_BEING_BUILT Then
                queueStructureBuildProgress psStruct, batchedMultiRectRenderer, 2
            End If
        End If
        it = lnNext(it)
    Loop
End Sub

Private Sub batchedObjectStatusRenderer_addStructureTargetingGfxToRender()
    ' Future todo: make this instanced as well?
    renderStructureTargetingGfx = True
End Sub

Private Sub batchedObjectStatusRenderer_drawAllSelections()
    Dim i As Long, k As Long, psDroid As Long, psStruct As Long, scrX As Long, scrY As Long, it As Long
    ' droid selection / health draw
    ' 1. box + power bars rects
    batchedMultiRectRenderer.drawRects True, 0
    ' 2. droid rank images
    For k = 0 To droidsToDraw_size - 1
        psDroid = droidsToDraw(k)
        ' Write the droid rank out
        If (aDroid(psDroid).sDisplay.screenX + aDroid(psDroid).sDisplay.screenR) > 0 _
           And (aDroid(psDroid).sDisplay.screenX - aDroid(psDroid).sDisplay.screenR) < pie_GetVideoBufferWidth() _
           And (aDroid(psDroid).sDisplay.screenY + aDroid(psDroid).sDisplay.screenR) > 0 _
           And (aDroid(psDroid).sDisplay.screenY - aDroid(psDroid).sDisplay.screenR) < pie_GetVideoBufferHeight() Then
            drawDroidRank psDroid
            drawDroidSensorLock psDroid
            drawDroidCmndNo psDroid
            drawGroupNumber psDroid
        End If
    Next
    ' 3. weapon reload bar
    batchedMultiRectRenderer.drawRects True, 1

    ' structure selection / health draw
    ' 1. structure health / build progress rects
    batchedMultiRectRenderer.drawRects True, 2
    ' 2. structure target origin icons
    For k = 0 To structsToDraw_size - 1
        psStruct = structsToDraw(k)
        For i = 0 To aStruct(psStruct).numWeaps - 1
            drawStructureTargetOriginIcon psStruct, i
        Next
        If STRUCTURE_isFactory(psStruct) Then
            drawGroupNumber psStruct
        End If
    Next
    ' 3. structure weapon reload bars
    batchedMultiRectRenderer.drawRects True, 3

    If renderStructureTargetingGfx Then
        For i = 0 To MAX_PLAYERS - 1
            it = apsStructLists(i).head
            Do While it <> 0
                psStruct = lnObj(it)
                ' If it's targetted and on-screen
                If (aStruct(psStruct).flags And BIT_(OBJECT_FLAG_TARGETED)) <> 0 _
                   And aStruct(psStruct).sDisplay.frameNumber = currentGameFrame Then
                    scrX = aStruct(psStruct).sDisplay.screenX
                    scrY = aStruct(psStruct).sDisplay.screenY
                    iV_DrawImage IntImages, getTargettingGfx(), scrX, scrY
                End If
                it = lnNext(it)
            Loop
        Next
    End If

    ' draw last rects
    batchedMultiRectRenderer.drawRects True, 4
End Sub

' BatchedObjectStatusRenderer::initialize (display3d_c's init3DView)
Public Sub batchedObjectStatusRenderer_initialize()
    If batchedMultiRectRenderer Is Nothing Then Set batchedMultiRectRenderer = New BatchedMultiRectRenderer
    batchedMultiRectRenderer.initialize
    ' 2 for droid selection rect groups, + 2 for structure selection rect groups, + 1 for "mouse over" rect groups (drawn last)
    batchedMultiRectRenderer.resizeRectGroups 5
End Sub

Public Sub batchedObjectStatusRenderer_clear()
    droidsToDraw_size = 0
    structsToDraw_size = 0
    renderStructureTargetingGfx = False
    If Not batchedMultiRectRenderer Is Nothing Then batchedMultiRectRenderer.clear
End Sub

Public Sub batchedObjectStatusRenderer_reset()
    batchedObjectStatusRenderer_clear
    If Not batchedMultiRectRenderer Is Nothing Then batchedMultiRectRenderer.reset
End Sub

Public Sub drawDragBox()
    Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, color As Long
    If dragBox3D.status <> DRAG_DRAGGING Or buildState <> BUILD3D_NONE Then
        Exit Sub
    End If

    If dragBox3D.x1 < mouseX() Then X1 = dragBox3D.x1 Else X1 = mouseX()
    If dragBox3D.x1 > mouseX() Then X2 = dragBox3D.x1 Else X2 = mouseX()
    If dragBox3D.y1 < mouseY() Then Y1 = dragBox3D.y1 Else Y1 = mouseY()
    If dragBox3D.y1 > mouseY() Then Y2 = dragBox3D.y1 Else Y2 = mouseY()

    ' draw static box

    iV_Box X1, Y1, X2, Y2, psPalette(WZCOL_UNIT_SELECT_BORDER)
    pie_UniTransBoxFill X1, Y1, X2, Y2, psPalette(WZCOL_UNIT_SELECT_BOX)

    ' draw pulse effect

    dragBox3D.pulse = dragBox3D.pulse + (BOX_PULSE_SIZE - dragBox3D.pulse) * realTimeAdjustedIncrement(5)

    If dragBox3D.pulse > BOX_PULSE_SIZE - 0.1! Then
        dragBox3D.pulse = 0
    End If

    color = psPalette(WZCOL_UNIT_SELECT_BOX)

    ' alpha relative to max pulse size
    color = (color And &HFFFFFF) Or SHL(Fix(CSng(SHRU(color, 24)) * (1 - (dragBox3D.pulse / BOX_PULSE_SIZE))) And &HFF&, 24)

    pie_UniTransBoxFill X2, Y1, X2 + dragBox3D.pulse, Y2 + dragBox3D.pulse, color ' east side + south-east corner
    pie_UniTransBoxFill X1 - dragBox3D.pulse, Y2, X2, Y2 + dragBox3D.pulse, color ' south side + south-west corner
    pie_UniTransBoxFill X1 - dragBox3D.pulse, Y1 - dragBox3D.pulse, X1, Y2, color ' west side + north-west corner
    pie_UniTransBoxFill X1, Y1 - dragBox3D.pulse, X2 + dragBox3D.pulse, Y1, color ' north side + north-east corner
End Sub

' Display reload bars for structures and droids
' psWeap: (VB) the object's weapon in weapon_slot (psObj->asWeaps[weapon_slot])
Private Sub queueWeaponReloadBarRects(ByVal psObj As Long, ByVal weapon_slot As Long, ByVal batched As BatchedMultiRectRenderer, ByVal rectGroup As Long)
    Dim scrX As Long, scrY As Long, scrR As Long, scale_ As Long, psStruct As Long, mulH As Single, psDroid As Long
    Dim armed As Long, firingStage As Long, bw As Long, bb As Long

    If ctrlShiftDown() And (SIMPLE_OBJECT_type(psObj) = OBJ_DROID) Then
        psDroid = psObj
        scrX = aDroid(psDroid).sDisplay.screenX
        scrY = aDroid(psDroid).sDisplay.screenY
        scrR = aDroid(psDroid).sDisplay.screenR
        scrY = scrY + scrR + 2

        If weapon_slot <> 0 Then ' only rendering resistance in the first slot
            Exit Sub
        End If
        If aDroid(psDroid).resistance <> 0 Then
            mulH = CSng(PERCENT_(aDroid(psDroid).resistance, DROID_droidResistance(psDroid)))
        Else
            mulH = 100!
        End If
        firingStage = Fix(mulH)
        firingStage = ((((2 * scrR) * 10000) \ 100) * firingStage) \ 10000
        If U32(firingStage) >= U32(2 * scrR) Then
            firingStage = (2 * scrR)
        End If
        batched.addRect scrX - scrR - 1, 3 + scrY + 0 + (weapon_slot * 2), scrX - scrR + (2 * scrR) + 1, 3 + scrY + 3 + (weapon_slot * 2), psPalette(WZCOL_RELOAD_BACKGROUND)
        batched.addRect scrX - scrR, 3 + scrY + 1 + (weapon_slot * 2), scrX - scrR + firingStage, 3 + scrY + 2 + (weapon_slot * 2), psPalette(WZCOL_HEALTH_RESISTANCE)
        Exit Sub
    End If

    If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        armed = droidReloadBar(psObj, aDroid(psObj).asWeaps(weapon_slot), weapon_slot)
    Else
        armed = droidReloadBar(psObj, aStruct(psObj).asWeaps(weapon_slot), weapon_slot)
    End If
    If armed >= 0 And armed < 100 Then ' no need to draw if full
        Select Case SIMPLE_OBJECT_type(psObj)
        Case OBJ_DROID
            scrX = aDroid(psObj).sDisplay.screenX
            scrY = aDroid(psObj).sDisplay.screenY
            scrR = aDroid(psObj).sDisplay.screenR
            scrY = scrY + scrR + 2
        Case OBJ_STRUCTURE
            psStruct = psObj
            scrX = aStruct(psStruct).sDisplay.screenX
            scrY = aStruct(psStruct).sDisplay.screenY
            scrR = aStruct(psStruct).sDisplay.screenR
            bw = asStructureStats(aStruct(psStruct).pStructureType).baseWidth
            bb = asStructureStats(aStruct(psStruct).pStructureType).baseBreadth
            If bw > bb Then scale_ = bw Else scale_ = bb
            scrY = scrY + scale_ * 10
            scrR = scale_ * 20
        End Select
        ' Scale it into an appropriate range
        firingStage = ((((2 * scrR) * 10000) \ 100) * armed) \ 10000
        If firingStage >= 2 * scrR Then
            firingStage = (2 * scrR)
        End If
        ' Power bars
        batched.addRect scrX - scrR - 1, 3 + scrY + 0 + (weapon_slot * 2), scrX - scrR + (2 * scrR) + 1, 3 + scrY + 3 + (weapon_slot * 2), psPalette(WZCOL_RELOAD_BACKGROUND), rectGroup
        batched.addRect scrX - scrR, 3 + scrY + 1 + (weapon_slot * 2), scrX - scrR + firingStage, 3 + scrY + 2 + (weapon_slot * 2), psPalette(WZCOL_RELOAD_BAR), rectGroup
    End If
End Sub

' draw target origin icon for the specified structure
Private Sub drawStructureTargetOriginIcon(ByVal psStruct As Long, ByVal weapon_slot As Long)
    Dim scrX As Long, scrY As Long, scrR As Long, scale_ As Long, bw As Long, bb As Long

    ' Process main weapon only for now
    If Not tuiTargetOrigin Or weapon_slot <> 0 Then
        Exit Sub
    End If
    If aStruct(psStruct).asWeaps(weapon_slot).nStat = 0 Then
        Exit Sub
    End If

    bw = asStructureStats(aStruct(psStruct).pStructureType).baseWidth
    bb = asStructureStats(aStruct(psStruct).pStructureType).baseBreadth
    If bw > bb Then scale_ = bw Else scale_ = bb
    scrX = aStruct(psStruct).sDisplay.screenX
    scrY = aStruct(psStruct).sDisplay.screenY + (scale_ * 10)
    scrR = scale_ * 20

    ' Render target origin graphics
    Select Case aStruct(psStruct).asWeaps(weapon_slot).origin
    Case ORIGIN_VISUAL
        iV_DrawImage IntImages, IMAGE_ORIGIN_VISUAL, scrX + scrR + 5, scrY - 1
    Case ORIGIN_COMMANDER
        iV_DrawImage IntImages, IMAGE_ORIGIN_COMMANDER, scrX + scrR + 5, scrY - 1
    Case ORIGIN_SENSOR
        iV_DrawImage IntImages, IMAGE_ORIGIN_SENSOR_STANDARD, scrX + scrR + 5, scrY - 1
    Case ORIGIN_CB_SENSOR
        iV_DrawImage IntImages, IMAGE_ORIGIN_SENSOR_CB, scrX + scrR + 5, scrY - 1
    Case ORIGIN_AIRDEF_SENSOR
        iV_DrawImage IntImages, IMAGE_ORIGIN_SENSOR_AIRDEF, scrX + scrR + 5, scrY - 1
    Case ORIGIN_RADAR_DETECTOR
        iV_DrawImage IntImages, IMAGE_ORIGIN_RADAR_DETECTOR, scrX + scrR + 5, scrY - 1
    Case ORIGIN_UNKNOWN
        ' Do nothing
    Case Else
        debug_ 3318, LOG_WARNING, "drawStructureTargetOriginIcon", "Unexpected target origin in structure(%d)!", aStruct(psStruct).id
    End Select
End Sub

' draw the health bar for the specified structure
Private Sub queueStructureHealth(ByVal psStruct As Long, ByVal batched As BatchedMultiRectRenderer, ByVal rectGroup As Long)
    Dim scrX As Long, scrY As Long, scrR As Long, powerCol As Long, powerColShadow As Long, health As Long, width As Long
    Dim scale_ As Long, bw As Long, bb As Long, resistance As Long, maxBody As Long, maxBodyBuilt As Long, r As Long
    powerCol = psPalette(WZCOL_BLACK): powerColShadow = psPalette(WZCOL_BLACK)

    bw = asStructureStats(aStruct(psStruct).pStructureType).baseWidth
    bb = asStructureStats(aStruct(psStruct).pStructureType).baseBreadth
    If bw > bb Then scale_ = bw Else scale_ = bb
    width = scale_ * 20
    scrX = aStruct(psStruct).sDisplay.screenX
    scrY = aStruct(psStruct).sDisplay.screenY + (scale_ * 10)
    scrR = width
    'health = PERCENT(psStruct->body, psStruct->baseBodyPoints);
    If ctrlShiftDown() Then
        'show resistance values if CTRL/SHIFT depressed
        resistance = structureResistance(aStruct(psStruct).pStructureType, aStruct(psStruct).player)
        If resistance <> 0 Then
            r = aStruct(psStruct).resistance
            If r < 0 Then r = 0
            health = PERCENT_(r, resistance)
        Else
            health = 100
        End If
    Else
        'show body points
        health = Fix((1# - getStructureDamage(psStruct) / 65536!) * 100)

        ' If structure is incomplete, make bar correspondingly thinner.
        maxBody = STRUCTURE_structureBody(psStruct)
        maxBodyBuilt = structureBodyBuilt(psStruct)
        width = Fix(CDbl(width) * maxBodyBuilt / maxBody)
    End If
    If health > REPAIRLEV_HIGH Then
        powerCol = psPalette(WZCOL_HEALTH_HIGH)
        powerColShadow = psPalette(WZCOL_HEALTH_HIGH_SHADOW)
    ElseIf health > REPAIRLEV_LOW Then
        powerCol = psPalette(WZCOL_HEALTH_MEDIUM)
        powerColShadow = psPalette(WZCOL_HEALTH_MEDIUM_SHADOW)
    Else
        powerCol = psPalette(WZCOL_HEALTH_LOW)
        powerColShadow = psPalette(WZCOL_HEALTH_LOW_SHADOW)
    End If
    health = (((width * 10000) \ 100) * health) \ 10000
    health = health * 2
    batched.addRectF scrX - scrR - 1, scrY - 1, scrX - scrR + 2 * width + 1, scrY + 3, psPalette(WZCOL_RELOAD_BACKGROUND), rectGroup
    batched.addRectF scrX - scrR, scrY, scrX - scrR + health, scrY + 1, powerCol, rectGroup
    batched.addRectF scrX - scrR, scrY + 1, scrX - scrR + health, scrY + 2, powerColShadow, rectGroup
End Sub

' draw the construction bar for the specified structure
Private Sub queueStructureBuildProgress(ByVal psStruct As Long, ByVal batched As BatchedMultiRectRenderer, ByVal rectGroup As Long)
    Dim scale_ As Long, scrX As Long, scrY As Long, scrR As Long, progress As Single, bw As Long, bb As Long
    bw = asStructureStats(aStruct(psStruct).pStructureType).baseWidth
    bb = asStructureStats(aStruct(psStruct).pStructureType).baseBreadth
    If bw > bb Then scale_ = bw Else scale_ = bb
    scrX = aStruct(psStruct).sDisplay.screenX
    scrY = aStruct(psStruct).sDisplay.screenY + (scale_ * 10)
    scrR = scale_ * 20
    progress = scale_ * 40 * structureCompletionProgress(psStruct)
    batched.addRectF scrX - scrR - 1, scrY - 1 + 5, scrX + scrR + 1, scrY + 3 + 5, psPalette(WZCOL_RELOAD_BACKGROUND), rectGroup
    batched.addRectF scrX - scrR, scrY + 5, scrX - scrR + progress, scrY + 1 + 5, psPalette(WZCOL_HEALTH_MEDIUM_SHADOW), rectGroup
    batched.addRectF scrX - scrR, scrY + 1 + 5, scrX - scrR + progress, scrY + 2 + 5, psPalette(WZCOL_HEALTH_MEDIUM), rectGroup
End Sub

' Draw the health of structures and show enemy structures being targetted
Private Sub drawStructureSelections()
    Dim psStruct As Long, psClickedOn As Long, bMouseOverStructure As Boolean, bMouseOverOwnStructure As Boolean
    Dim isSpectator As Boolean, player As Long, btn As Long

    psClickedOn = mouseTarget()
    If psClickedOn <> 0 Then
        If SIMPLE_OBJECT_type(psClickedOn) = OBJ_STRUCTURE Then
            bMouseOverStructure = True
            If SIMPLE_OBJECT_player(psClickedOn) = selectedPlayer Then
                bMouseOverOwnStructure = True
            End If
        End If
    End If
    pie_SetFogStatus False
    If bMultiPlayer Then isSpectator = NetPlay_players_isSpectator(selectedPlayer)
    If isSpectator Then
        If barMode = BAR_DROIDS_AND_STRUCTURES Then
            For player = 0 To MAX_PLAYERS - 1
                batchedObjectStatusRenderer_addStructureSelectionsToRender player, psClickedOn, bMouseOverOwnStructure
            Next
        End If
        Exit Sub
    End If
    If selectedPlayer >= MAX_PLAYERS Then Exit Sub ' no-op

    batchedObjectStatusRenderer_addStructureSelectionsToRender selectedPlayer, psClickedOn, bMouseOverOwnStructure

    batchedObjectStatusRenderer_addStructureTargetingGfxToRender

    If bMouseOverStructure And Not bMouseOverOwnStructure Then
        If getRightClickOrders() Then btn = MOUSE_LMB Else btn = MOUSE_RMB
        If mouseDown(btn) Then
            psStruct = psClickedOn
            batchedObjectStatusRenderer_addEnemyStructureHealthToRender psStruct
        End If
    End If
End Sub

Private Function getTargettingGfx() As Long
    Dim index As Long
    index = getModularScaledRealTime(1000, 10)
    Select Case index
    Case 0, 1, 2
        getTargettingGfx = (IMAGE_TARGET1 + index)
    Case Else
        If (index And &H1) <> 0 Then
            getTargettingGfx = (IMAGE_TARGET4)
        Else
            getTargettingGfx = (IMAGE_TARGET5)
        End If
    End Select
End Function

' Is the droid, its commander or its sensor tower selected?
Public Function eitherSelected(ByVal psDroid As Long) As Boolean
    Dim psObj As Long
    If aDroid(psDroid).selected Then
        eitherSelected = True
        Exit Function
    End If

    If aDroid(psDroid).psGroup <> 0 Then
        If aGroup(aDroid(psDroid).psGroup).psCommander <> 0 Then
            If aDroid(aGroup(aDroid(psDroid).psGroup).psCommander).selected Then
                eitherSelected = True
                Exit Function
            End If
        End If
    End If

    psObj = orderStateObj(psDroid, DORDER_FIRESUPPORT)
    If psObj <> 0 Then
        If BASE_OBJECT_selected(psObj) Then
            eitherSelected = True
            Exit Function
        End If
    End If

    eitherSelected = False
End Function

Private Sub queueDroidPowerBarsRects(ByVal psDroid As Long, ByVal drawBox As Boolean, ByVal batched As BatchedMultiRectRenderer, ByVal rectGroup As Long)
    Dim damage As Long, shields As Long, maxShieldPoints As Long, powerCol As Long, powerColShadow As Long
    Dim sx As Long, sy As Long, sr As Long, white As Long
    damage = PERCENT_(aDroid(psDroid).body, aDroid(psDroid).originalBody)
    shields = 0

    If aDroid(psDroid).shieldPoints >= 0 Then
        maxShieldPoints = droidGetMaxShieldPoints(psDroid)

        If maxShieldPoints > 0 Then
            shields = PERCENT_(aDroid(psDroid).shieldPoints, maxShieldPoints)
            shields = Fix(CSng(aDroid(psDroid).shieldPoints) / CSng(maxShieldPoints) * CSng(aDroid(psDroid).sDisplay.screenR))
        End If
    End If

    powerCol = psPalette(WZCOL_BLACK)
    powerColShadow = psPalette(WZCOL_BLACK)

    If U32(damage) > REPAIRLEV_HIGH Then
        powerCol = psPalette(WZCOL_HEALTH_HIGH)
        powerColShadow = psPalette(WZCOL_HEALTH_HIGH_SHADOW)
    ElseIf U32(damage) > REPAIRLEV_LOW Then
        powerCol = psPalette(WZCOL_HEALTH_MEDIUM)
        powerColShadow = psPalette(WZCOL_HEALTH_MEDIUM_SHADOW)
    Else
        powerCol = psPalette(WZCOL_HEALTH_LOW)
        powerColShadow = psPalette(WZCOL_HEALTH_LOW_SHADOW)
    End If

    damage = Fix(CSng(aDroid(psDroid).body) / CSng(aDroid(psDroid).originalBody) * CSng(aDroid(psDroid).sDisplay.screenR))
    If U32(damage) > U32(aDroid(psDroid).sDisplay.screenR) Then
        damage = aDroid(psDroid).sDisplay.screenR
    End If

    damage = damage * 2

    sx = aDroid(psDroid).sDisplay.screenX
    sy = aDroid(psDroid).sDisplay.screenY
    sr = aDroid(psDroid).sDisplay.screenR
    white = psPalette(WZCOL_WHITE)
    If drawBox Then
        batched.addRect sx - sr, sy + sr - 7, sx - sr + 1, sy + sr, white, rectGroup
        batched.addRect sx - sr, sy + sr, sx - sr + 7, sy + sr + 1, white, rectGroup
        batched.addRect sx + sr - 7, sy + sr, sx + sr, sy + sr + 1, white, rectGroup
        batched.addRect sx + sr, sy + sr - 7, sx + sr + 1, sy + sr + 1, white, rectGroup
    End If

    ' Power bars
    batched.addRect sx - sr - 1, sy + sr + 2, sx + sr + 1, sy + sr + 6, psPalette(WZCOL_RELOAD_BACKGROUND), rectGroup
    batched.addRect sx - sr, sy + sr + 3, sx - sr + damage, sy + sr + 4, powerCol, rectGroup
    batched.addRect sx - sr, sy + sr + 4, sx - sr + damage, sy + sr + 5, powerColShadow, rectGroup
    batched.addRect sx - sr - 3, sy + sr, sx - sr - 1, sy + sr - shields, psPalette(WZCOL_LBLUE), rectGroup
    batched.addRect sx - sr - 2, sy + sr, sx - sr, sy + sr - shields, psPalette(WZCOL_BLACK), rectGroup
End Sub

Private Sub queueDroidEnemyHealthBarsRects(ByVal psDroid As Long, ByVal batched As BatchedMultiRectRenderer, ByVal rectGroup As Long)
    Dim powerCol As Long, powerColShadow As Long, boxCol As Long, damage As Long, mulH As Single
    Dim scrX As Long, scrY As Long, scrR As Long
    powerCol = psPalette(WZCOL_BLACK)
    powerColShadow = psPalette(WZCOL_BLACK)

    'show resistance values if CTRL/SHIFT depressed
    If ctrlShiftDown() Then
        If aDroid(psDroid).resistance <> 0 Then
            damage = PERCENT_(aDroid(psDroid).resistance, DROID_droidResistance(psDroid))
        Else
            damage = 100
        End If
    Else
        damage = PERCENT_(aDroid(psDroid).body, aDroid(psDroid).originalBody)
    End If

    If U32(damage) > REPAIRLEV_HIGH Then
        powerCol = psPalette(WZCOL_HEALTH_HIGH)
        powerColShadow = psPalette(WZCOL_HEALTH_HIGH_SHADOW)
    ElseIf U32(damage) > REPAIRLEV_LOW Then
        powerCol = psPalette(WZCOL_HEALTH_MEDIUM)
        powerColShadow = psPalette(WZCOL_HEALTH_MEDIUM_SHADOW)
    Else
        powerCol = psPalette(WZCOL_HEALTH_LOW)
        powerColShadow = psPalette(WZCOL_HEALTH_LOW_SHADOW)
    End If

    'show resistance values if CTRL/SHIFT depressed
    If ctrlShiftDown() Then
        If aDroid(psDroid).resistance <> 0 Then
            mulH = CSng(aDroid(psDroid).resistance) / CSng(DROID_droidResistance(psDroid))
        Else
            mulH = 100!
        End If
    Else
        mulH = CSng(aDroid(psDroid).body) / CSng(aDroid(psDroid).originalBody)
    End If
    damage = Fix(mulH * CSng(aDroid(psDroid).sDisplay.screenR)) ' (((psDroid->sDisplay.screenR*10000)/100)*damage)/10000;
    If U32(damage) > U32(aDroid(psDroid).sDisplay.screenR) Then
        damage = aDroid(psDroid).sDisplay.screenR
    End If
    damage = damage * 2
    scrX = aDroid(psDroid).sDisplay.screenX
    scrY = aDroid(psDroid).sDisplay.screenY
    scrR = aDroid(psDroid).sDisplay.screenR

    ' Three DFX clips properly right now - not sure if software does
    If (scrX + scrR) > 0 _
       And (scrX - scrR) < pie_GetVideoBufferWidth() _
       And (scrY + scrR) > 0 _
       And (scrY - scrR) < pie_GetVideoBufferHeight() Then
        boxCol = psPalette(WZCOL_WHITE)

        ' Power bars
        batched.addRect scrX - scrR - 1, scrY + scrR + 2, scrX + scrR + 1, scrY + scrR + 6, psPalette(WZCOL_RELOAD_BACKGROUND), rectGroup
        batched.addRect scrX - scrR, scrY + scrR + 3, scrX - scrR + damage, scrY + scrR + 4, powerCol, rectGroup
        batched.addRect scrX - scrR, scrY + scrR + 4, scrX - scrR + damage, scrY + scrR + 5, powerColShadow, rectGroup
    End If
End Sub

' Draw the selection graphics for selected droids
Private Sub drawDroidSelections()
    Dim psClickedOn As Long, bMouseOverDroid As Boolean, bMouseOverOwnDroid As Boolean, index As Long
    Dim isSpectator As Boolean, player As Long, i As Long, it As Long, psDroid As Long, psFeature As Long, btn As Long

    psClickedOn = mouseTarget()
    If psClickedOn <> 0 Then
        If SIMPLE_OBJECT_type(psClickedOn) = OBJ_DROID Then
            bMouseOverDroid = True
            If SIMPLE_OBJECT_player(psClickedOn) = selectedPlayer And Not BASE_OBJECT_selected(psClickedOn) Then
                bMouseOverOwnDroid = True
            End If
        End If
    End If
    If bMultiPlayer Then isSpectator = NetPlay_players_isSpectator(selectedPlayer)

    If isSpectator Then
        If barMode = BAR_DROIDS Or barMode = BAR_DROIDS_AND_STRUCTURES Then
            For player = 0 To MAX_PLAYERS - 1
                batchedObjectStatusRenderer_addDroidSelectionsToRender player, psClickedOn, bMouseOverDroid
            Next
        End If

        Exit Sub
    End If

    If selectedPlayer >= MAX_PLAYERS Then Exit Sub ' no-op
    pie_SetFogStatus False

    batchedObjectStatusRenderer_addDroidSelectionsToRender selectedPlayer, psClickedOn, bMouseOverDroid

    ' Are we over an enemy droid
    If bMouseOverDroid And Not bMouseOverOwnDroid Then
        If getRightClickOrders() Then btn = MOUSE_LMB Else btn = MOUSE_RMB
        If mouseDown(btn) Then
            If SIMPLE_OBJECT_player(psClickedOn) <> selectedPlayer And aDroid(psClickedOn).sDisplay.frameNumber = currentGameFrame Then
                psDroid = psClickedOn
                batchedObjectStatusRenderer_addEnemyDroidHealthToRender psDroid
            End If
        End If
    End If

    For i = 0 To MAX_PLAYERS - 1
        ' Go thru' all the droids
        it = apsDroidLists(i).head
        Do While it <> 0
            psDroid = lnObj(it)
            If showORDERS Then
                drawDroidOrder psDroid
            End If
            If aDroid(psDroid).died = 0 And aDroid(psDroid).sDisplay.frameNumber = currentGameFrame Then
                ' If it's selected
                If (aDroid(psDroid).flags And BIT_(OBJECT_FLAG_TARGETED)) <> 0 And BASE_OBJECT_visibleForLocalDisplay(psDroid) = UBYTE_MAX Then
                    index = IMAGE_BLUE1 + getModularScaledRealTime(1020, 5)
                    iV_DrawImage IntImages, index, aDroid(psDroid).sDisplay.screenX, aDroid(psDroid).sDisplay.screenY
                End If
            End If
            it = lnNext(it)
        Loop
    Next

    it = apsFeatureList(0).head
    Do While it <> 0
        psFeature = lnObj(it)
        If aFeature(psFeature).died = 0 And aFeature(psFeature).sDisplay.frameNumber = currentGameFrame Then
            If (aFeature(psFeature).flags And BIT_(OBJECT_FLAG_TARGETED)) <> 0 Then
                iV_DrawImage IntImages, getTargettingGfx(), aFeature(psFeature).sDisplay.screenX, aFeature(psFeature).sDisplay.screenY
            End If
        End If
        it = lnNext(it)
    Loop
End Sub

Public Sub drawDroidAndStructureSelections()
    drawDroidSelections
    drawStructureSelections

    batchedObjectStatusRenderer_drawAllSelections
    batchedObjectStatusRenderer_clear
End Sub

' rendering of the object's group next to the object itself,
' or the group that will be assigned to the object after production in the factory
Private Sub drawGroupNumber(ByVal psObject As Long)
    Dim id As Long, groupNumber As Long, x As Long, y As Long, groupNumberType As GROUPNUMBER_TYPE
    Dim psDroid As Long, psStruct As Long, xShift As Long, yShift As Long, scale_ As Long, width As Long
    Dim scrX As Long, scrY As Long, scrR As Long, bw As Long, bb As Long
    id = UWORD_MAX
    groupNumber = UBYTE_MAX
    x = 0: y = 0
    groupNumberType = GN_NORMAL

    psDroid = castDroid(psObject)
    psStruct = castStructure(psObject)
    If psDroid <> 0 Then
        xShift = aDroid(psDroid).sDisplay.screenR + GN_X_OFFSET
        yShift = aDroid(psDroid).sDisplay.screenR

        x = aDroid(psDroid).sDisplay.screenX - xShift
        y = aDroid(psDroid).sDisplay.screenY + yShift

        If aDroid(psDroid).repairGroup <> UBYTE_MAX Then
            groupNumber = aDroid(psDroid).repairGroup
            groupNumberType = GN_DAMAGED
        Else
            groupNumber = aDroid(psDroid).group
        End If
    ElseIf psStruct <> 0 Then
        ' same as in queueStructureHealth
        bw = asStructureStats(aStruct(psStruct).pStructureType).baseWidth
        bb = asStructureStats(aStruct(psStruct).pStructureType).baseBreadth
        If bw > bb Then scale_ = bw Else scale_ = bb
        width = scale_ * 20
        scrX = aStruct(psStruct).sDisplay.screenX
        scrY = aStruct(psStruct).sDisplay.screenY + (scale_ * 10)
        scrR = width

        x = scrX - scrR - GN_X_OFFSET
        y = scrY - GN_Y_OFFSET

        groupNumber = aStruct(psStruct).productToGroup
        groupNumberType = GN_FACTORY
    End If

    Select Case groupNumber
    Case 0
        id = IMAGE_GN_0
    Case 1
        id = IMAGE_GN_1
    Case 2
        id = IMAGE_GN_2
    Case 3
        id = IMAGE_GN_3
    Case 4
        id = IMAGE_GN_4
    Case 5
        id = IMAGE_GN_5
    Case 6
        id = IMAGE_GN_6
    Case 7
        id = IMAGE_GN_7
    Case 8
        id = IMAGE_GN_8
    Case 9
        id = IMAGE_GN_9
    End Select

    If id <> UWORD_MAX Then
        Select Case groupNumberType
        Case GN_NORMAL
            iV_DrawImage IntImages, id, x, y
        Case GN_DAMAGED
            iV_DrawImageTint IntImages, id, x, y, pal_RGBA(255, 0, 0, 255) ' red
        Case GN_FACTORY
            iV_DrawImageTint IntImages, id, x, y, pal_RGBA(255, 220, 115, 255) ' gold
        End Select
    End If
End Sub

Private Sub drawDroidOrder(ByVal psDroid As Long)
    Dim xShift As Long, yShift As Long, letter As String
    xShift = aDroid(psDroid).sDisplay.screenR + GN_X_OFFSET
    yShift = aDroid(psDroid).sDisplay.screenR - CMND_GN_Y_OFFSET
    letter = getDroidOrderKey(aDroid(psDroid).order.typ)
    iV_SetTextColour psPalette(WZCOL_TEXT_BRIGHT)
    iV_DrawText letter, aDroid(psDroid).sDisplay.screenX - xShift - CMND_STAR_X_OFFSET, aDroid(psDroid).sDisplay.screenY + yShift, font_regular
End Sub

' Draw the number of the commander the droid is assigned to
Private Sub drawDroidCmndNo(ByVal psDroid As Long)
    Dim xShift As Long, yShift As Long, index As Long, id2 As Long, id As Long, bDraw As Boolean
    bDraw = True

    id = UWORD_MAX

    id2 = IMAGE_GN_STAR
    index = SDWORD_MAX
    If aDroid(psDroid).droidType = DROID_COMMAND Then
        index = cmdDroidGetIndex(psDroid)
    ElseIf hasCommander(psDroid) Then
        index = cmdDroidGetIndex(aGroup(aDroid(psDroid).psGroup).psCommander)
    End If
    Select Case index
    Case 1
        id = IMAGE_GN_1
    Case 2
        id = IMAGE_GN_2
    Case 3
        id = IMAGE_GN_3
    Case 4
        id = IMAGE_GN_4
    Case 5
        id = IMAGE_GN_5
    Case 6
        id = IMAGE_GN_6
    Case 7
        id = IMAGE_GN_7
    Case 8
        id = IMAGE_GN_8
    Case 9
        id = IMAGE_GN_9
    Case Else
        bDraw = False
    End Select

    If bDraw Then
        xShift = aDroid(psDroid).sDisplay.screenR + GN_X_OFFSET
        yShift = aDroid(psDroid).sDisplay.screenR - CMND_GN_Y_OFFSET
        iV_DrawImage IntImages, id2, aDroid(psDroid).sDisplay.screenX - xShift - CMND_STAR_X_OFFSET, aDroid(psDroid).sDisplay.screenY + yShift
        iV_DrawImage IntImages, id, aDroid(psDroid).sDisplay.screenX - xShift, aDroid(psDroid).sDisplay.screenY + yShift
    End If
End Sub

Public Function toggleEnergyBars() As ENERGY_BAR
    barMode = barMode + 1
    If barMode = BAR_LAST Then
        barMode = BAR_SELECTED
    End If
    toggleEnergyBars = barMode
End Function

' Set everything up for when the player assigns the sensor target
Public Sub assignSensorTarget(ByVal psObj As Long)
    bSensorTargetting = True
    lastTargetAssignation = realTime
    psSensorObj = psObj
End Sub

' Set everything up for when the player selects the destination
Public Sub assignDestTarget()
    bDestTargetting = True
    lastDestAssignation = realTime
    destTargetX = mouseX()
    destTargetY = mouseY()
    destTileX = mouseTileX
    destTileY = mouseTileY
End Sub

' Draw a graphical effect after selecting a sensor target
Public Sub processSensorTarget()
    Dim x As Long, y As Long, index As Long, offset As Long, x0 As Long, y0 As Long, x1 As Long, y1 As Long
    Dim lines(0 To 9) As Vector4i, showIt As Boolean
    If bSensorTargetting Then
        If U32(realTime - lastTargetAssignation) < TARGET_TO_SENSOR_TIME Then
            showIt = (SIMPLE_OBJECT_died(psSensorObj) = 0)
            If showIt Then showIt = (BASE_OBJECT_sDisplay_frameNumber(psSensorObj) = currentGameFrame)
            If showIt Then
                x = ToInt16(BASE_OBJECT_sDisplay_screenX(psSensorObj)) ' /*mouseX();*/(SWORD)
                y = ToInt16(BASE_OBJECT_sDisplay_screenY(psSensorObj))
                index = IMAGE_BLUE1
                If Not gamePaused() Then
                    index = IMAGE_BLUE1 + getModularScaledGraphicsTime(1020, 5)
                End If
                iV_DrawImage IntImages, index, x, y
                offset = 12 + ((TARGET_TO_SENSOR_TIME) - (realTime - lastTargetAssignation)) \ 2
                x0 = x - offset
                y0 = y - offset
                x1 = x + offset
                y1 = y + offset
                lines(0) = glm_ivec4(x0, y0, x0 + 8, y0): lines(1) = glm_ivec4(x0, y0, x0, y0 + 8)
                lines(2) = glm_ivec4(x1, y0, x1 - 8, y0): lines(3) = glm_ivec4(x1, y0, x1, y0 + 8)
                lines(4) = glm_ivec4(x1, y1, x1 - 8, y1): lines(5) = glm_ivec4(x1, y1, x1, y1 - 8)
                lines(6) = glm_ivec4(x0, y1, x0 + 8, y1): lines(7) = glm_ivec4(x0, y1, x0, y1 - 8)
                lines(8) = glm_ivec4(x0, y0, x0 + 8, y0): lines(9) = glm_ivec4(x0, y0, x0, y0 + 8)
                iV_Lines VarPtr(lines(0)), 10, psPalette(WZCOL_WHITE)
            Else
                bSensorTargetting = False
            End If
        Else
            bSensorTargetting = False
        End If
    End If
End Sub

' Draw a graphical effect after selecting a destination
Public Sub processDestinationTarget()
    Dim x As Long, y As Long, offset As Long, x0 As Long, y0 As Long, x1 As Long, y1 As Long
    If bDestTargetting Then
        If U32(realTime - lastDestAssignation) < DEST_TARGET_TIME Then
            x = destTargetX
            y = destTargetY
            offset = ((DEST_TARGET_TIME) - (realTime - lastDestAssignation)) \ 2
            x0 = x - offset
            y0 = y - offset
            x1 = x + offset
            y1 = y + offset

            pie_BoxFill x0, y0, x0 + 2, y0 + 2, psPalette(WZCOL_WHITE)
            pie_BoxFill x1 - 2, y0 - 2, x1, y0, psPalette(WZCOL_WHITE)
            pie_BoxFill x1 - 2, y1 - 2, x1, y1, psPalette(WZCOL_WHITE)
            pie_BoxFill x0, y1, x0 + 2, y1 + 2, psPalette(WZCOL_WHITE)
        Else
            bDestTargetting = False
        End If
    End If
End Sub

' Set what tile is being used to draw the bottom of a body of water
Public Sub setUnderwaterTile(ByVal num As Long)
    underwaterTile = num
End Sub

' Set what tile is being used to show rubble
Public Sub setRubbleTile(ByVal num As Long)
    rubbleTile = num
End Sub

' Get the tile that is currently being used to draw underwater ground
Public Function getWaterTileNum() As Long
    getWaterTileNum = underwaterTile
End Function

' Get the tile that is being used to show rubble
Public Function getRubbleTileNum() As Long
    getRubbleTileNum = rubbleTile
End Function

' Draw the spinning particles for power stations and re-arm pads for the specified player
Private Sub structureEffectsPlayer(ByVal player As Long)
    Dim radius As Long, xDif As Long, yDif As Long, pos As Vector3i, gameDiv As Long
    Dim effectsPerSecond As Long, effectTime As Long, it As Long, psStructure As Long, numConnected As Long, i As Long
    Dim psChosenObj As Long, bFXSize As Long, psDroid As Long, imd As Long, mx As Long, my As Long, mz As Long, typ As Long

    effectsPerSecond = 12  ' Effects per second. Will add effects up to once time per frame, so won't add as many effects if the framerate is low, but will be consistent, otherwise.
    effectTime = FromU32(Fix(U32(graphicsTime) / (GAME_TICKS_PER_SEC \ effectsPerSecond)) * (GAME_TICKS_PER_SEC \ effectsPerSecond))
    If U32(effectTime) <= U32(graphicsTime - deltaGraphicsTime) Then
        Exit Sub  ' Don't add effects this frame.
    End If

    it = apsStructLists(player).head
    Do While it <> 0
        psStructure = lnObj(it)
        If aStruct(psStructure).status <> SS_BUILT Then
            GoTo continue_
        End If
        typ = asStructureStats(aStruct(psStructure).pStructureType).typ
        If typ = REF_POWER_GEN And BASE_OBJECT_visibleForLocalDisplay(psStructure) <> 0 Then
            numConnected = 0
            For i = 0 To NUM_POWER_MODULES - 1
                If aStruct(psStructure).functionality.powerGenerator.apResExtractors(i) <> 0 Then
                    numConnected = numConnected + 1
                End If
            Next
            ' No effect if nothing connected
            If numConnected = 0 Then
                'keep looking for another!
                GoTo continue_
            Else
                Select Case numConnected
                Case 1, 2
                    gameDiv = 1440
                Case Else ' 3, 4
                    gameDiv = 1080   ' really fast!!!
                End Select
            End If

            ' New addition - it shows how many are connected...
            For i = 0 To numConnected - 1
                radius = 32 - (i * 2)    ' around the spire
                xDif = iSinSR(effectTime, gameDiv, radius)
                yDif = iCosSR(effectTime, gameDiv, radius)

                pos.x = aStruct(psStructure).pos.x + xDif
                pos.z = aStruct(psStructure).pos.y + yDif
                pos.y = map_Height(pos.x, pos.z) + 64 + (i * 20)   ' 64 up to get to base of spire
                effectGiveAuxVar 50    ' half normal plasma size...
                addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_LASER, False, 0, 0, graphicsTime

                pos.x = aStruct(psStructure).pos.x - xDif
                pos.z = aStruct(psStructure).pos.y - yDif
                effectGiveAuxVar 50    ' half normal plasma size...

                addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_LASER, False, 0, 0, graphicsTime
            Next
        ' Might be a re-arm pad!
        ElseIf typ = REF_REARM_PAD And BASE_OBJECT_visibleForLocalDisplay(psStructure) <> 0 Then
            psChosenObj = aStruct(psStructure).functionality.rearmPad.psObj
            If psChosenObj <> 0 Then
                If BASE_OBJECT_visibleForLocalDisplay(psChosenObj) <> 0 Then
                    bFXSize = 0
                    psDroid = psChosenObj
                    If aDroid(psDroid).died = 0 And aDroid(psDroid).action = DACTION_WAITDURINGREARM Then
                        bFXSize = 30
                    End If
                    ' Then it's repairing...?
                    imd = aStruct(psStructure).sDisplay.imd   ' the display model's radius and max are its base shape's
                    radius = iIMDBaseShape_radius(imd)
                    xDif = iSinSR(effectTime, 720, radius)
                    yDif = iCosSR(effectTime, 720, radius)
                    iIMDBaseShape_max imd, mx, my, mz
                    pos.x = aStruct(psStructure).pos.x + xDif
                    pos.z = aStruct(psStructure).pos.y + yDif
                    pos.y = map_Height(pos.x, pos.z) + my
                    effectGiveAuxVar 30 + bFXSize  ' half normal plasma size...
                    addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_LASER, False, 0, 0, graphicsTime
                    pos.x = aStruct(psStructure).pos.x - xDif
                    pos.z = aStruct(psStructure).pos.y - yDif   ' buildings are level!
                    effectGiveAuxVar 30 + bFXSize  ' half normal plasma size...
                    addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_LASER, False, 0, 0, graphicsTime
                End If
            End If
        End If
continue_:
        it = lnNext(it)
    Loop
End Sub

' Draw the effects for all players and buildings
Public Sub structureEffects()
    Dim i As Long
    For i = 0 To MAX_PLAYERS - 1
        If apsStructLists(i).size <> 0 Then
            structureEffectsPlayer i
        End If
    Next
End Sub

' Show the sensor ranges of selected droids and buildings
Public Sub showDroidSensorRanges()
    Static lastRangeUpdateTime As Long
    Dim it As Long

    If selectedPlayer >= MAX_PLAYERS Then Exit Sub ' no-op

    If rangeOnScreen And U32(graphicsTime - lastRangeUpdateTime) >= 50 Then ' note, we still have to decide what to do with multiple units selected, since it will draw it for all of them! -Q 5-10-05
        it = apsDroidLists(selectedPlayer).head
        Do While it <> 0
            If aDroid(lnObj(it)).selected Then
                showSensorRange2 lnObj(it)
            End If
            it = lnNext(it)
        Loop

        it = apsStructLists(selectedPlayer).head
        Do While it <> 0
            If aStruct(lnObj(it)).selected Then
                showSensorRange2 lnObj(it)
            End If
            it = lnNext(it)
        Loop

        lastRangeUpdateTime = graphicsTime
    End If 'end if we want to display...
End Sub

Private Sub showEffectCircle(centre As Vector3i, ByVal radius As Long, ByVal auxVar As Long, ByVal group As EFFECT_GROUP, ByVal typ As EFFECT_TYPE)
    Dim circumference As Long, i As Long, pos As Vector3i
    circumference = radius * 2 * 355 \ 113 \ TILE_UNITS ' 2 pi r in tiles.
    For i = 0 To circumference - 1
        pos.x = centre.x - iSinSR(i, circumference, radius)
        pos.z = centre.y - iCosSR(i, circumference, radius)  ' [sic] y -> z

        ' Check if it's actually on map
        If worldOnMap(pos.x, pos.z) Then
            pos.y = map_Height(pos.x, pos.z) + 16
            effectGiveAuxVar auxVar
            addEffect pos, group, typ, False, 0, 0, graphicsTime
        End If
    Next
End Sub

' Shows the weapon (long) range of the object in question.
' Note, it only does it for the first weapon slot!
Private Sub showWeaponRange(ByVal psObj As Long)
    Dim psStats As Long, compIndex As Long, psStruct As Long, weaponRange As Long, minRange As Long, pos As Vector3i

    If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        compIndex = aDroid(psObj).asWeaps(0).nStat   ' weapon_slot
        If Not (compIndex < asWeaponStats_size) Then ASSERT_FAILURE "compIndex < asWeaponStats.size()", "src/display3d.cpp:4397", "showWeaponRange", "Invalid range referenced for numWeaponStats, %d > %zu", compIndex, asWeaponStats_size: Exit Sub
        psStats = compIndex
    Else
        psStruct = psObj
        If asStructureStats(aStruct(psStruct).pStructureType).numWeaps = 0 Then
            Exit Sub
        End If
        psStats = asStructureStats(aStruct(psStruct).pStructureType).psWeapStat(0)
    End If
    weaponRange = proj_GetLongRange(asWeaponStats(psStats), SIMPLE_OBJECT_player(psObj))
    minRange = proj_GetMinRange(asWeaponStats(psStats), SIMPLE_OBJECT_player(psObj))
    pos = SIMPLE_OBJECT_pos(psObj)
    showEffectCircle pos, weaponRange, 40, EFFECT_EXPLOSION, EXPLOSION_TYPE_SMALL
    If minRange > 0 Then
        showEffectCircle pos, minRange, 40, EFFECT_EXPLOSION, EXPLOSION_TYPE_TESLA
    End If
End Sub

Private Sub showSensorRange2(ByVal psObj As Long)
    Dim pos As Vector3i
    pos = SIMPLE_OBJECT_pos(psObj)
    showEffectCircle pos, objSensorRange(psObj), 80, EFFECT_EXPLOSION, EXPLOSION_TYPE_LASER
    showWeaponRange psObj
End Sub

' Draw a circle on the map (to show the range of something)
Public Sub drawRangeAtPos(ByVal centerX As Long, ByVal centerY As Long, ByVal radius As Long)
    Dim pos As Vector3i
    pos.x = centerX: pos.y = centerY: pos.z = 0  ' .z ignored.
    showEffectCircle pos, radius, 80, EFFECT_EXPLOSION, EXPLOSION_TYPE_SMALL
End Sub

'* Turn on drawing some effects at certain position to visualize the radius.
' * \note Pass a negative radius to turn this off
Public Sub showRangeAtPos(ByVal centerX As Long, ByVal centerY As Long, ByVal radius As Long)
    rangeCenterX = centerX
    rangeCenterY = centerY
    rangeRadius = radius

    bRangeDisplay = True

    If radius <= 0 Then
        bRangeDisplay = False
    End If
End Sub

Public Function getDroidRankGraphicFromLevel(ByVal level As Long) As Long
    Dim gfxId As Long
    gfxId = UDWORD_MAX
    Select Case level
    Case 0
    Case 1
        gfxId = IMAGE_LEV_0
    Case 2
        gfxId = IMAGE_LEV_1
    Case 3
        gfxId = IMAGE_LEV_2
    Case 4
        gfxId = IMAGE_LEV_3
    Case 5
        gfxId = IMAGE_LEV_4
    Case 6
        gfxId = IMAGE_LEV_5
    Case 7
        gfxId = IMAGE_LEV_6
    Case 8
        gfxId = IMAGE_LEV_7
    Case Else
        ASSERT_FAILURE "!""out of range droid rank""", "src/display3d.cpp:4479", "getDroidRankGraphicFromLevel", "Weird droid level in drawDroidRank"
    End Select

    getDroidRankGraphicFromLevel = gfxId
End Function

' Get the graphic ID for a droid rank
Public Function getDroidRankGraphic(ByVal psDroid As Long) As Long
    ' Establish the numerical value of the droid's rank
    getDroidRankGraphic = getDroidRankGraphicFromLevel(getDroidLevel(psDroid))
End Function

'*    Will render a graphic depiction of the droid's present rank.
' * \note Assumes matrix context set and that z-buffer write is force enabled (Always).
Private Sub drawDroidRank(ByVal psDroid As Long)
    Dim gfxId As Long
    gfxId = getDroidRankGraphic(psDroid)

    ' Did we get one? - We should have...
    If gfxId <> UDWORD_MAX Then
        ' Render the rank graphic at the correct location */ // remove hardcoded numbers?!
        iV_DrawImage IntImages, gfxId And &HFFFF&, _
                     aDroid(psDroid).sDisplay.screenX + aDroid(psDroid).sDisplay.screenR + 8, _
                     aDroid(psDroid).sDisplay.screenY + aDroid(psDroid).sDisplay.screenR
    End If
End Sub

'*    Will render a sensor graphic for a droid locked to a sensor droid/structure
' * \note Assumes matrix context set and that z-buffer write is force enabled (Always).
Private Sub drawDroidSensorLock(ByVal psDroid As Long)
    'if on fire support duty - must be locked to a Sensor Droid/Structure
    If orderState(psDroid, DORDER_FIRESUPPORT) Then
        ' Render the sensor graphic at the correct location - which is what?!
        iV_DrawImage IntImages, IMAGE_GN_STAR, aDroid(psDroid).sDisplay.screenX, aDroid(psDroid).sDisplay.screenY
    End If
End Sub

' Draw the construction lines for all construction droids
Public Sub doConstructionLines(viewMatrix As mat4)
    Dim i As Long, it As Long, psDroid As Long, a As Long
    For i = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(i).head
        Do While it <> 0
            psDroid = lnObj(it)
            If clipXY(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y) Then
                If BASE_OBJECT_visibleForLocalDisplay(psDroid) = UBYTE_MAX And aDroid(psDroid).sMove.status <> MOVESHUFFLE Then
                    a = aDroid(psDroid).action
                    If a = DACTION_BUILD Then
                        If aDroid(psDroid).order.psObj <> 0 Then
                            If SIMPLE_OBJECT_type(aDroid(psDroid).order.psObj) = OBJ_STRUCTURE Then
                                addConstructionLine psDroid, aDroid(psDroid).order.psObj, viewMatrix
                            End If
                        End If
                    ElseIf (a = DACTION_DEMOLISH) Or _
                           (a = DACTION_REPAIR) Or _
                           (a = DACTION_RESTORE) Then
                        If aDroid(psDroid).psActionTarget(0) <> 0 Then
                            If SIMPLE_OBJECT_type(aDroid(psDroid).psActionTarget(0)) = OBJ_STRUCTURE Then
                                addConstructionLine psDroid, aDroid(psDroid).psActionTarget(0), viewMatrix
                            End If
                        End If
                    End If
                End If
            End If
            it = lnNext(it)
        Loop
    Next
End Sub

' (VB) randHash's lambda: v += d; v *= x; v ^= v>>15; v *= 0x987decaf; v ^= v>>17 (uint32_t)
Private Sub randHash_shuffle(v As Long, ByVal d As Long, ByVal x As Long)
    v = FromU32((U32(v) + U32(d)) - IIf(U32(v) + U32(d) >= 4294967296#, 4294967296#, 0#))
    v = MUL32(v, x)
    v = v Xor SHRU(v, 15)
    v = MUL32(v, &H987DECAF)
    v = v Xor SHRU(v, 17)
End Sub

' randHash(std::initializer_list<uint32_t> data): data(0 .. n - 1)
Private Function randHash(data() As Long, ByVal n As Long) As Long
    Dim v As Long, i As Long
    v = &H12345678
    For i = 0 To n - 1
        randHash_shuffle v, data(i), &H7EA99999
    Next
    For i = 0 To n - 1
        randHash_shuffle v, data(i), &HC0FFEE77
    Next
    randHash = v
End Function

' (VB) addConstructionLine's getPoint lambda: [&] captures amount, psDroid, psStructure, the model and each
Private Function addConstructionLine_getPoint(ByVal c As Long, ByVal amount As Long, ByVal psDroid As Long, ByVal psStructure As Long, ByVal pDisplayModel As Long, each_ As Vector3i, deltaPlayer As Vector3f) As Vector3f
    Dim t As Long, s As Single, data(0 To 4) As Long, npoints As Long, pointIndexA As Long, pointIndexB As Long
    Dim ax As Single, ay As Single, az As Single, bx As Single, by_ As Single, bz As Single, point_ As Vector3f, sum As Double, h As Double
    sum = U32(amount) + U32(c)
    If sum >= 4294967296# Then sum = sum - 4294967296#
    t = FromU32(Fix(sum / 1000))
    s = (sum - Fix(sum / 1000) * 1000) * 0.001!
    npoints = iIMDShape_points_size(pDisplayModel)
    data(0) = aDroid(psDroid).id: data(1) = aStruct(psStructure).id: data(2) = aDroid(psDroid).actionStarted: data(3) = t: data(4) = c
    h = U32(randHash(data, 5))
    pointIndexA = FromU32(h - Fix(h / npoints) * npoints)
    data(3) = t + 1
    h = U32(randHash(data, 5))
    pointIndexB = FromU32(h - Fix(h / npoints) * npoints)
    iIMDShape_point pDisplayModel, pointIndexA, ax, ay, az
    iIMDShape_point pDisplayModel, pointIndexB, bx, by_, bz
    ' auto point = mix(pointA, pointB, s)
    point_.x = ax + (bx - ax) * s
    point_.y = ay + (by_ - ay) * s
    point_.z = az + (bz - az) * s

    ' each (a Vector3i) = Vector3f(pos.x, pos.z, pos.y) + Vector3f(point_.x, structHeightScale(psStructure) * point_.y, -point_.z)
    each_.x = Fix(aStruct(psStructure).pos.x + point_.x)
    each_.y = Fix(aStruct(psStructure).pos.z + structHeightScale(psStructure) * point_.y)
    each_.z = Fix(aStruct(psStructure).pos.y + -point_.z)
    addConstructionLine_getPoint.x = each_.x + deltaPlayer.x
    addConstructionLine_getPoint.y = each_.y + deltaPlayer.y
    addConstructionLine_getPoint.z = -each_.z + deltaPlayer.z
End Function

' Draw the construction or demolish lines for one droid
Private Sub addConstructionLine(ByVal psDroid As Long, ByVal psStructure As Long, viewMatrix As mat4)
    Dim pDisplayModel As Long, deltaPlayer As Vector3f, pt(0 To 2) As Vector3f, tri(0 To 2) As Vector3f
    Dim constructPoints As Long, amount As Long, each_ As Vector3i, pointC As Vector3f, cross As Vector3f, shift As Vector3f
    Dim f As Single, colour As Long, a As Vector3f
    If aStruct(psStructure).sDisplay.imd = 0 Then
        Exit Sub
    End If
    pDisplayModel = iIMDBaseShape_displayModel(aStruct(psStructure).sDisplay.imd)

    deltaPlayer = glm_vec3(0, 0, 0)
    pt(0).x = aDroid(psDroid).pos.x + deltaPlayer.x
    pt(0).y = aDroid(psDroid).pos.z + 24 + deltaPlayer.y
    pt(0).z = -aDroid(psDroid).pos.y + deltaPlayer.z

    constructPoints = constructorPoints(asConstructStats(DROID_getConstructStats(psDroid)), aDroid(psDroid).player)
    ' int amount = 800 * constructPoints * (graphicsTime - psDroid->actionStarted) / GAME_TICKS_PER_SEC (unsigned)
    amount = FromU32(Fix(U32(MUL32(800 * constructPoints, graphicsTime - aDroid(psDroid).actionStarted)) / GAME_TICKS_PER_SEC))

    pt(1) = addConstructionLine_getPoint(250, amount, psDroid, psStructure, pDisplayModel, each_, deltaPlayer)
    pt(2) = addConstructionLine_getPoint(750, amount, psDroid, psStructure, pDisplayModel, each_, deltaPlayer)

    If aStruct(psStructure).currentBuildPts < 10 Then
        pointC.x = aStruct(psStructure).pos.x + deltaPlayer.x
        pointC.y = aStruct(psStructure).pos.z + 10 + deltaPlayer.y
        pointC.z = -aStruct(psStructure).pos.y + deltaPlayer.z
        cross = glm_vec3(aStruct(psStructure).pos.y - aDroid(psDroid).pos.y, 0, aStruct(psStructure).pos.x - aDroid(psDroid).pos.x)
        shift = glm_vec3_mul(glm_normalize3(cross), 40!)
        f = aStruct(psStructure).currentBuildPts * 0.1!
        ' pt1 = mix(pointC - shift, pt1, f); pt2 = mix(pointC + shift, pt1, f)
        a = glm_vec3_sub(pointC, shift)
        tri(1).x = a.x + (pt(1).x - a.x) * f: tri(1).y = a.y + (pt(1).y - a.y) * f: tri(1).z = a.z + (pt(1).z - a.z) * f
        a = glm_vec3_add(pointC, shift)
        tri(2).x = a.x + (pt(1).x - a.x) * f: tri(2).y = a.y + (pt(1).y - a.y) * f: tri(2).z = a.z + (pt(1).z - a.z) * f
        pt(1) = tri(1)
        pt(2) = tri(2)
    End If

    If U32(rand() Mod 250) < U32(deltaGraphicsTime) Then
        effectSetSize 30
        addEffect each_, EFFECT_EXPLOSION, EXPLOSION_TYPE_SPECIFIED, True, getDisplayImdFromIndex(MI_PLASMA), 0, graphicsTime
    End If

    If aDroid(psDroid).action = DACTION_DEMOLISH Then colour = psPalette(WZCOL_DEMOLISH_BEAM) Else colour = psPalette(WZCOL_CONSTRUCTOR_BEAM)
    tri(0) = pt(0): tri(1) = pt(1): tri(2) = pt(2)
    pie_TransColouredTriangle VarPtr(tri(0)), colour, VarPtr(viewMatrix)
    tri(0) = pt(0): tri(1) = pt(2): tri(2) = pt(1)
    pie_TransColouredTriangle VarPtr(tri(0)), colour, VarPtr(viewMatrix)
End Sub

Public Function getDrawTerrainShadows() As Boolean
    getDrawTerrainShadows = bDrawTerrainShadows
End Function

Public Sub setDrawTerrainShadows(ByVal val_ As Boolean)
    bDrawTerrainShadows = val_
End Sub
