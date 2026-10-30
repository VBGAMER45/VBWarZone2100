Attribute VB_Name = "multistruct_c"
' src/multistruct.cpp: multiplayer structure related stuff: the debug
' messages that add and remove a structure, the lassat firing, and
' GAME_STRUCTUREINFO (production, research hold and release). The local
' DROID_TEMPLATE of recvStructureInfo is a UDT; only a manufacture order
' uses it, through the pool copy copyTemplate makes (pT). structure.h's
' static inline popStatusPending is a private copy per functionality type,
' as structure_c has its setStatusPending* copies.
Option Explicit

' weapondef.h
Private Const HEAVY_DAMAGE_LEVEL As Long = 25

' structure.h's popStatusPending, for a FACTORY
Private Sub popStatusPending_factory(func_ As FACTORY)
    If func_.pendingCount = 0 Then
        func_.pendingCount = func_.pendingCount + 1
    End If
    func_.pendingCount = func_.pendingCount - 1
    If func_.pendingCount = 0 Then
        ' Subject is now synchronised, remove pending.
        func_.psSubjectPending = 0
        func_.statusPending = FACTORY_NOTHING_PENDING
    End If
End Sub

' structure.h's popStatusPending, for a RESEARCH_FACILITY
Private Sub popStatusPending_research(func_ As RESEARCH_FACILITY)
    If func_.pendingCount = 0 Then
        func_.pendingCount = func_.pendingCount + 1
    End If
    func_.pendingCount = func_.pendingCount - 1
    If func_.pendingCount = 0 Then
        ' Subject is now synchronised, remove pending.
        func_.psSubjectPending = -1
        func_.statusPending = FACTORY_NOTHING_PENDING
    End If
End Sub

' ////////////////////////////////////////////////////////////////////////////
' structures

' ////////////////////////////////////////////////////////////////////////////
' INFORM others that a building has been completed.
Public Function SendBuildFinished(ByVal psStruct As Long) As Boolean
    Dim player As Long, w As MessageWriter, id As Long, ref As Long, pos As Vector3i

    player = aStruct(psStruct).player
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/multistruct.cpp:64", "SendBuildFinished", "invalid player %u", player: SendBuildFinished = False: Exit Function

    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_DEBUG_ADD_STRUCTURE)
    id = aStruct(psStruct).id
    NETuint32_t w, id                   ' ID of building

    ' Along with enough info to build it (if needed)
    ref = asStructureStats(aStruct(psStruct).pStructureType).ref
    NETuint32_t w, ref
    pos = aStruct(psStruct).pos
    NETPosition w, pos
    NETuint8_t w, player
    SendBuildFinished = NETend(w)
End Function

' ////////////////////////////////////////////////////////////////////////////
Public Function recvBuildFinished(ByVal queue As Long) As Boolean
    Dim structId As Long, psStruct As Long, pos As Vector3i, typ As Long, typeindex As Long, player As Long
    Dim r As MessageReader

    Set r = NETbeginDecode(queue, GAME_DEBUG_ADD_STRUCTURE)
    NETuint32_t r, structId     ' get the struct id.
    NETuint32_t r, typ          ' Kind of building.
    NETPosition r, pos          ' pos
    NETuint8_t r, player
    NETend r

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/multistruct.cpp:92", "recvBuildFinished", "invalid player %u", player: recvBuildFinished = False: Exit Function

    If Not debugMappingsAllowed() And bMultiPlayer Then
        debug_ 97, LOG_WARNING, "recvBuildFinished", "Failed to add structure for player %u.", NetPlay_players_position(NETQUEUE_index(queue))
        recvBuildFinished = False
        Exit Function
    End If

    psStruct = IdToStruct(structId, ANYPLAYER)

    If psStruct <> 0 Then
        ' make it complete.
        aStruct(psStruct).currentBuildPts = structureBuildPointsToCompletion(psStruct) + 1

        If aStruct(psStruct).status <> SS_BUILT Then
            debug_ 110, LOG_SYNC, "recvBuildFinished", "Synch error, structure %u was not complete, and should have been.", structId
            aStruct(psStruct).status = SS_BUILT
            buildingComplete psStruct
        End If
        debug_ 114, LOG_SYNC, "recvBuildFinished", "Created normal building %u for player %u", aStruct(psStruct).id, player
        recvBuildFinished = True
        Exit Function
    End If

    ' The building wasn't started, so we'll have to just plonk it down in the map.

    ' Find the structures stats
    typeindex = 0
    Do While typeindex < numStructureStats
        If asStructureStats(typeindex).ref = typ Then Exit Do
        typeindex = typeindex + 1
    Loop    ' Find structure target

    ' Build the structure
    psStruct = buildStructureDir(typeindex, pos.x, pos.y, 0, player, True, structId)
    If psStruct <> 0 Then
        aStruct(psStruct).status = SS_BUILT
        buildingComplete psStruct
        debug_ 129, LOG_SYNC, "recvBuildFinished", "Huge synch error, forced to create building %u for player %u", aStruct(psStruct).id, player
        ' (NETlogEntry "had to plonk down a building": #if defined (DEBUG))
        triggerEventStructBuilt psStruct, 0
        checkPlayerBuiltHQ psStruct
    Else
        debug_ 138, LOG_SYNC, "recvBuildFinished", "Gigantic synch error, unable to create building for player %u", player
        NETlogEntry "had to plonk down a building, BUT FAILED!", SYNC_FLAG, player
    End If

    recvBuildFinished = False
End Function


' ////////////////////////////////////////////////////////////////////////////
' Inform others that a structure has been destroyed
Public Function SendDestroyStructure(ByVal s As Long) As Boolean
    Dim w As MessageWriter, id As Long

    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_DEBUG_REMOVE_STRUCTURE)
    ' Struct to destroy
    id = aStruct(s).id
    NETuint32_t w, id

    SendDestroyStructure = NETend(w)
End Function

' ////////////////////////////////////////////////////////////////////////////
' acknowledge the destruction of a structure, from another player.
Public Function recvDestroyStructure(ByVal queue As Long) As Boolean
    Dim structID As Long, psStruct As Long, r As MessageReader

    Set r = NETbeginDecode(queue, GAME_DEBUG_REMOVE_STRUCTURE)
    NETuint32_t r, structID
    NETend r

    If Not debugMappingsAllowed() And bMultiPlayer Then
        debug_ 171, LOG_WARNING, "recvDestroyStructure", "Failed to remove structure for player %u.", NetPlay_players_position(NETQUEUE_index(queue))
        recvDestroyStructure = False
        Exit Function
    End If

    ' Struct to destroy
    psStruct = IdToStruct(structID, ANYPLAYER)

    If psStruct <> 0 Then
        turnOffMultiMsg True
        ' Remove the struct from remote players machine
        destroyStruct psStruct, FromU32(CDbl(gameTime) - deltaGameTime + 1)  ' deltaGameTime is actually 0 here, since we're between updates. However, the value of gameTime - deltaGameTime + 1 will not change when we start the next tick.
        turnOffMultiMsg False
    End If

    recvDestroyStructure = True
End Function

' ////////////////////////////////////////////////////////////////////////////
'lassat is firing

Public Function sendLasSat(ByVal player As Long, ByVal psStruct As Long, ByVal psObj As Long) As Boolean
    Dim w As MessageWriter, id As Long, targetId As Long, targetPlayer As Long

    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_LASSAT)
    NETuint8_t w, player
    id = aStruct(psStruct).id
    NETuint32_t w, id
    targetId = SIMPLE_OBJECT_id(psObj)
    NETuint32_t w, targetId             ' Target
    targetPlayer = SIMPLE_OBJECT_player(psObj)
    NETuint8_t w, targetPlayer          ' Target player

    sendLasSat = NETend(w)
End Function

' recv lassat info on the receiving end.
Public Function recvLasSat(ByVal queue As Long) As Boolean
    Dim psObj As Long, player As Long, targetplayer As Long, psStruct As Long, id As Long, targetid As Long
    Dim r As MessageReader, firePause As Long, damLevel As Long, target As Vector3i

    Set r = NETbeginDecode(queue, GAME_LASSAT)
    NETuint8_t r, player
    NETuint32_t r, id
    NETuint32_t r, targetid
    NETuint8_t r, targetplayer
    NETend r

    psStruct = IdToStruct(id, player)
    psObj = IdToPointer(targetid, targetplayer)
    If psStruct <> 0 Then
        If Not canGiveOrdersFor(NETQUEUE_index(queue), aStruct(psStruct).player) Then
            syncDebug "recvLasSat", "Wrong player."
            recvLasSat = False
            Exit Function
        End If
    End If

    If psStruct <> 0 And psObj <> 0 Then
        If isLasSat(aStruct(psStruct).pStructureType) Then
            ' Lassats have just one weapon
            firePause = weaponFirePause(asWeaponStats(BASE_OBJECT_getWeaponStats(psStruct, 0)), player)
            damLevel = FromU32(Fix(U32(MUL32(aStruct(psStruct).body, 100)) / U32(STRUCTURE_structureBody(psStruct))))  ' PERCENT(psStruct->body, psStruct->structureBody())

            If ULT(damLevel, HEAVY_DAMAGE_LEVEL) Then
                firePause = FromU32(U32(firePause) + U32(firePause))
            End If

            If isHumanPlayer(player) Then
                If Not ULT(firePause, FromU32(U32(gameTime) - U32(aStruct(psStruct).asWeaps(0).lastFired))) Then
                    ' Too soon to fire again
                    recvLasSat = True Xor False  ' Return value meaningless and ignored.
                    Exit Function
                End If
            End If

            ' Give enemy no quarter, unleash the lasat
            target = SIMPLE_OBJECT_pos(psObj)
            proj_SendProjectile aStruct(psStruct).asWeaps(0), 0, player, target, psObj, True, 0
            aStruct(psStruct).asWeaps(0).lastFired = gameTime
            aStruct(psStruct).asWeaps(0).ammo = 1 ' abducting this field for keeping track of triggers

            ' Play 5 second countdown message
            audio_QueueTrackPos ID_SOUND_LAS_SAT_COUNTDOWN, target.x, target.y, target.z
        End If
    End If

    recvLasSat = True
End Function

' (VB) pT: an apsTemplates handle, used only when structureInfo_ is STRUCTUREINFO_MANUFACTURE
Public Sub sendStructureInfo(ByVal psStruct As Long, ByVal structureInfo_ As STRUCTURE_INFO, ByVal pT As Long)
    Dim player As Long, structId As Long, structureInfo As Long, w As MessageWriter, droidType As Long, i As Long

    player = aStruct(psStruct).player
    structId = aStruct(psStruct).id
    structureInfo = structureInfo_ And &HFF&

    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_STRUCTUREINFO)
    NETuint8_t w, player
    NETuint32_t w, structId
    NETuint8_t w, structureInfo
    If structureInfo_ = STRUCTUREINFO_MANUFACTURE Then
        droidType = apsTemplates(pT).droidType
        NETwzstring w, apsTemplates(pT).name_
        NETuint32_t w, apsTemplates(pT).multiPlayerID
        NETint32_t w, droidType
        NETuint8_t w, apsTemplates(pT).asParts(COMP_BODY)
        NETuint8_t w, apsTemplates(pT).asParts(COMP_BRAIN)
        NETuint8_t w, apsTemplates(pT).asParts(COMP_PROPULSION)
        NETuint8_t w, apsTemplates(pT).asParts(COMP_REPAIRUNIT)
        NETuint8_t w, apsTemplates(pT).asParts(COMP_ECM)
        NETuint8_t w, apsTemplates(pT).asParts(COMP_SENSOR)
        NETuint8_t w, apsTemplates(pT).asParts(COMP_CONSTRUCT)
        NETint8_t w, apsTemplates(pT).numWeaps
        For i = 0 To apsTemplates(pT).numWeaps - 1
            NETuint32_t w, apsTemplates(pT).asWeaps(i)
        Next
    End If
    NETend w
End Sub

Public Sub recvStructureInfo(ByVal queue As Long)
    Dim player As Long, structId As Long, structureInfo As Long, psStruct As Long
    Dim t As DROID_TEMPLATE, pT As Long, droidType As Long, r As MessageReader, name_ As String, i As Long
    Dim ch As Long

    player = 0
    structId = 0
    DROID_TEMPLATE_construct t
    pT = 0                              ' (VB) &t: only a manufacture order uses pT, after copyTemplate

    Set r = NETbeginDecode(queue, GAME_STRUCTUREINFO)
    NETuint8_t r, player
    NETuint32_t r, structId
    NETuint8_t r, structureInfo
    If structureInfo = STRUCTUREINFO_MANUFACTURE Then
        NETwzstring r, name_
        t.name_ = name_
        NETuint32_t r, t.multiPlayerID
        NETint32_t r, droidType
        NETuint8_t r, t.asParts(COMP_BODY)
        NETuint8_t r, t.asParts(COMP_BRAIN)
        NETuint8_t r, t.asParts(COMP_PROPULSION)
        NETuint8_t r, t.asParts(COMP_REPAIRUNIT)
        NETuint8_t r, t.asParts(COMP_ECM)
        NETuint8_t r, t.asParts(COMP_SENSOR)
        NETuint8_t r, t.asParts(COMP_CONSTRUCT)
        NETint8_t r, t.numWeaps
        If Not (t.numWeaps >= 0 And t.numWeaps <= MAX_WEAPONS) Then ASSERT_FAILURE "pT->numWeaps >= 0 && pT->numWeaps <= ARRAY_SIZE(pT->asWeaps)", "src/multistruct.cpp:315", "recvStructureInfo", "Bad numWeaps %d", t.numWeaps: Exit Sub
        For i = 0 To t.numWeaps - 1
            NETuint32_t r, t.asWeaps(i)
        Next
        t.droidType = droidType
        pT = copyTemplate(player, t)
    End If
    NETend r

    psStruct = IdToStruct(structId, player)

    If psStruct = 0 Then ch = Asc("^") Else ch = Asc("*")
    syncDebug "recvStructureInfo", "player%d,structId%u%c,structureInfo%u", player, structId, ch, structureInfo

    If psStruct = 0 Then
        debug_ 331, LOG_WARNING, "recvStructureInfo", "Could not find structure %u to change production for", structId
        Exit Sub
    End If
    If Not canGiveOrdersFor(NETQUEUE_index(queue), aStruct(psStruct).player) Then
        syncDebug "recvStructureInfo", "Wrong player."
        Exit Sub
    End If

    checkStructure psStruct, "src/multistruct.cpp:340", "recvStructureInfo", max_check_object_recursion

    If structureInfo = STRUCTUREINFO_MANUFACTURE Then
        If Not researchedTemplate(pT, player, True, True) Then
            debug_ 344, LOG_ERROR, "recvStructureInfo", "Invalid droid received from player %d with name %s", player, apsTemplates(pT).name_
            Exit Sub
        End If
    End If
    If structureInfo = STRUCTUREINFO_MANUFACTURE Then
        If Not intValidTemplate(apsTemplates(pT), vbNullString, True, player) Then
            debug_ 349, LOG_ERROR, "recvStructureInfo", "Illegal droid received from player %d with name %s", player, apsTemplates(pT).name_
            Exit Sub
        End If
    End If

    If STRUCTURE_isFactory(psStruct) Then
        popStatusPending_factory aStruct(psStruct).functionality.factory
    ElseIf asStructureStats(aStruct(psStruct).pStructureType).typ = REF_RESEARCH Then
        popStatusPending_research aStruct(psStruct).functionality.researchFacility
    End If

    syncDebugStructure_ "recvStructureInfo", psStruct, Asc("<")

    Select Case structureInfo
    Case STRUCTUREINFO_MANUFACTURE:       structSetManufacture psStruct, pT, ModeImmediate
    Case STRUCTUREINFO_CANCELPRODUCTION:  cancelProduction psStruct, ModeImmediate, False
    Case STRUCTUREINFO_HOLDPRODUCTION:    holdProduction psStruct, ModeImmediate
    Case STRUCTUREINFO_RELEASEPRODUCTION: releaseProduction psStruct, ModeImmediate
    Case STRUCTUREINFO_HOLDRESEARCH:      holdResearch psStruct, ModeImmediate
    Case STRUCTUREINFO_RELEASERESEARCH:   releaseResearch psStruct, ModeImmediate
    Case Else
        debug_ 373, LOG_ERROR, "recvStructureInfo", "Invalid structureInfo %d", structureInfo
    End Select

    syncDebugStructure_ "recvStructureInfo", psStruct, Asc(">")

    checkStructure psStruct, "src/multistruct.cpp:378", "recvStructureInfo", max_check_object_recursion
End Sub
