Attribute VB_Name = "multibot_c"
' src/multibot.cpp: multiplay stuff relevant to droids only: the droid
' orders queued by sendDroidInfo and sendDroidSecondary and sent, grouped,
' as GAME_DROIDINFO messages by sendQueuedDroidInfo, and their receipt; the
' debug messages that add and remove a droid; a droid leaving a
' transporter. The template NETQueuedDroidInfo<SerdeContext> is one copy
' taking the MessageReader or MessageWriter As Object (WzNet's serde
' procedures take either). std::map<DROID_ORDER, std::vector<QueuedDroidInfo>>
' in sendQueuedDroidInfo is the orders' values in ascending order, each
' group in insertion order. TargetMissing (processDroidTarget's error
' return) is a handle in no pool (OBJ_NUM_TYPES * OBJ_HANDLE_UNIT): never 0,
' never an object, and the SIMPLE_OBJECT_* accessors read nothing from it,
' as C++'s TargetMissing_ is never referenced. The local DROID_TEMPLATE of
' recvDroid is a template of the pool (reallyBuildDroid takes a handle),
' deleted when the function returns.
Option Explicit

Private Enum SubType
    ObjOrder
    LocOrder
    SecondaryOrder
End Enum

' (the fields serialised ByRef are Longs: WzNet's procedures take Long values)
Private Type QueuedDroidInfo
    player As Long                      ' uint8_t
    droidId As Long                     ' uint32_t
    subType As Long                     ' SubType
    ' subType == ObjOrder || subType == LocOrder
    order As Long                       ' DROID_ORDER
    destId As Long                      ' uint32_t: if (subType == ObjOrder)
    destType As Long                    ' OBJECT_TYPE: if (subType == ObjOrder)
    pos As Vector2i                     ' if (subType == LocOrder)
    y As Long                           ' uint32_t: if (subType == LocOrder)
    structRef As Long                   ' uint32_t: if (order == DORDER_BUILD || order == DORDER_LINEBUILD)
    direction As Long                   ' uint16_t: if (order == DORDER_BUILD || order == DORDER_LINEBUILD)
    index As Long                       ' uint32_t: if (order == DORDER_BUILDMODULE)
    pos2 As Vector2i                    ' if (order == DORDER_LINEBUILD)
    add As Boolean
    ' subType == SecondaryOrder
    secOrder As Long                    ' SECONDARY_ORDER (DSO_UNUSED)
    secState As Long                    ' SECONDARY_STATE (DSS_NONE)
End Type

Private queuedOrders() As QueuedDroidInfo
Private queuedOrders_size As Long

' processDroidTarget's error return value (see the header)
Private Const TargetMissing As Long = OBJ_NUM_TYPES * OBJ_HANDLE_UNIT

' (VB) a C comparison of two values: -1, 0 or 1 for <, ==, >
Private Function cmp3(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then cmp3 = -1 Else cmp3 = 1
End Function

' (VB) cmp3 for uint32_t values
Private Function cmp3U(ByVal a As Long, ByVal b As Long) As Long
    If ULT(a, b) Then cmp3U = -1 Else cmp3U = 1
End Function

' Returns 0 if order is the same, non-zero otherwise.
Private Function QueuedDroidInfo_orderCompare(this As QueuedDroidInfo, z As QueuedDroidInfo) As Long
    If this.player <> z.player Then
        QueuedDroidInfo_orderCompare = cmp3(this.player, z.player)
        Exit Function
    End If
    If this.subType <> z.subType Then
        QueuedDroidInfo_orderCompare = cmp3(this.subType, z.subType)
        Exit Function
    End If
    Select Case this.subType
    Case ObjOrder, LocOrder
        If this.order <> z.order Then
            QueuedDroidInfo_orderCompare = cmp3(this.order, z.order)
            Exit Function
        End If
        If this.subType = ObjOrder Then
            If this.destId <> z.destId Then
                QueuedDroidInfo_orderCompare = cmp3U(this.destId, z.destId)
                Exit Function
            End If
            If this.destType <> z.destType Then
                QueuedDroidInfo_orderCompare = cmp3(this.destType, z.destType)
                Exit Function
            End If
        Else
            If this.pos.x <> z.pos.x Then
                QueuedDroidInfo_orderCompare = cmp3(this.pos.x, z.pos.x)
                Exit Function
            End If
            If this.pos.y <> z.pos.y Then
                QueuedDroidInfo_orderCompare = cmp3(this.pos.y, z.pos.y)
                Exit Function
            End If
        End If
        If this.order = DORDER_BUILD Or this.order = DORDER_LINEBUILD Then
            If this.structRef <> z.structRef Then
                QueuedDroidInfo_orderCompare = cmp3U(this.structRef, z.structRef)
                Exit Function
            End If
            If this.direction <> z.direction Then
                QueuedDroidInfo_orderCompare = cmp3(this.direction, z.direction)
                Exit Function
            End If
        End If
        If this.order = DORDER_LINEBUILD Then
            If this.pos2.x <> z.pos2.x Then
                QueuedDroidInfo_orderCompare = cmp3(this.pos2.x, z.pos2.x)
                Exit Function
            End If
            If this.pos2.y <> z.pos2.y Then
                QueuedDroidInfo_orderCompare = cmp3(this.pos2.y, z.pos2.y)
                Exit Function
            End If
        End If
        If this.order = DORDER_BUILDMODULE Then
            If this.index <> z.index Then
                QueuedDroidInfo_orderCompare = cmp3U(this.index, z.index)
                Exit Function
            End If
        End If
        If this.add <> z.add Then
            If Not this.add Then QueuedDroidInfo_orderCompare = -1 Else QueuedDroidInfo_orderCompare = 1   ' add < z.add
            Exit Function
        End If
    Case SecondaryOrder
        If this.secOrder <> z.secOrder Then
            QueuedDroidInfo_orderCompare = cmp3(this.secOrder, z.secOrder)
            Exit Function
        End If
        If this.secState <> z.secState Then
            QueuedDroidInfo_orderCompare = cmp3(this.secState, z.secState)
            Exit Function
        End If
    End Select
    QueuedDroidInfo_orderCompare = 0
End Function

' Sorts by order, then finally by droid id, to group multiple droids with the same order. (operator <)
Private Function QueuedDroidInfo_less(this As QueuedDroidInfo, z As QueuedDroidInfo) As Boolean
    Dim orComp As Long
    orComp = QueuedDroidInfo_orderCompare(this, z)
    If orComp <> 0 Then
        QueuedDroidInfo_less = orComp < 0
        Exit Function
    End If
    QueuedDroidInfo_less = ULT(this.droidId, z.droidId)
End Function

' (VB) queuedOrders.push_back(info)
Private Sub queuedOrders_push_back(info As QueuedDroidInfo)
    If queuedOrders_size = 0 Then
        ReDim queuedOrders(0 To 15)
    ElseIf queuedOrders_size > UBound(queuedOrders) Then
        ReDim Preserve queuedOrders(0 To 2 * queuedOrders_size - 1)
    End If
    queuedOrders(queuedOrders_size) = info
    queuedOrders_size = queuedOrders_size + 1
End Sub

' ////////////////////////////////////////////////////////////////////////////
' Command Droids.

' sod em.


' ////////////////////////////////////////////////////////////////////////////
' Secondary Orders.

' Send
Public Function sendDroidSecondary(ByVal psDroid As Long, ByVal sec As SECONDARY_ORDER, ByVal state As SECONDARY_STATE) As Boolean
    Dim info As QueuedDroidInfo

    If Not bMultiMessages Then
        sendDroidSecondary = True
        Exit Function
    End If

    info.player = aDroid(psDroid).player
    info.droidId = aDroid(psDroid).id
    info.subType = SecondaryOrder
    info.secOrder = sec
    info.secState = state

    ' Send later, grouped by order, so multiple droids with the same order can be encoded to much less data.
    queuedOrders_push_back info

    sendDroidSecondary = True
End Function

' Broadcast that droid is being unloaded from a transporter
'
'  \sa recvDroidDisEmbark()
Public Function sendDroidDisembark(ByVal psTransporter As Long, ByVal psDroid As Long) As Boolean
    Dim w As MessageWriter, player As Long, droidId As Long, transportId As Long

    If Not bMultiMessages Then
        sendDroidDisembark = True
        Exit Function
    End If

    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_DROIDDISEMBARK)
    player = aDroid(psTransporter).player
    droidId = aDroid(psDroid).id
    transportId = aDroid(psTransporter).id

    NETuint32_t w, player
    NETuint32_t w, droidId
    NETuint32_t w, transportId

    sendDroidDisembark = NETend(w)
End Function

' Receive info about a droid that is being unloaded from a transporter
Public Function recvDroidDisEmbark(ByVal queue As Long) As Boolean
    Dim psFoundDroid As Long, psTransporterDroid As Long, r As MessageReader
    Dim player As Long, droidID As Long, transporterID As Long, it As Long

    psFoundDroid = 0
    psTransporterDroid = 0

    Set r = NETbeginDecode(queue, GAME_DROIDDISEMBARK)

    NETuint32_t r, player
    NETuint32_t r, droidID
    NETuint32_t r, transporterID

    NETend r

    ' find the transporter first
    psTransporterDroid = IdToDroid(transporterID, player)
    If psTransporterDroid = 0 Then
        ' Possible it already died? (sync error?)
        debug_ 263, LOG_WARNING, "recvDroidDisEmbark", "player's %d transport droid %d wasn't found?", player, transporterID
        recvDroidDisEmbark = False
        Exit Function
    End If
    If Not canGiveOrdersFor(NETQUEUE_index(queue), aDroid(psTransporterDroid).player) Then
        recvDroidDisEmbark = False
        Exit Function
    End If
    ' we need to find the droid *in* the transporter
    If aDroid(psTransporterDroid).psGroup <> 0 Then
        it = aGroup(aDroid(psTransporterDroid).psGroup).psList.head
        Do While it <> 0
            If aDroid(lnObj(it)).id = droidID Then
                psFoundDroid = lnObj(it)
                Exit Do
            End If
            it = lnNext(it)
        Loop
    End If
    ' don't continue if we couldn't find it.
    If psFoundDroid = 0 Then
        ' I don't think this could ever be possible...but
        debug_ 288, LOG_ERROR, "recvDroidDisEmbark", "Couldn't find droid %d to disembark from player %d's transporter?", droidID, player
        recvDroidDisEmbark = False
        Exit Function
    End If

    transporterRemoveDroid psTransporterDroid, psFoundDroid, ModeImmediate
    recvDroidDisEmbark = True
End Function


' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' Droids

' ////////////////////////////////////////////////////////////////////////////
' Send a new Droid to the other players
' (VB) initialOrdersP is hasInitialOrders (initialOrdersP != nullptr) and initialOrders (*initialOrdersP)
Public Function SendDroid(ByVal pTemplate As Long, ByVal x As Long, ByVal y As Long, ByVal player As Long, ByVal id As Long, ByVal hasInitialOrders As Boolean, initialOrders As INITIAL_DROID_ORDERS) As Boolean
    Dim w As MessageWriter, pos As Vector3i, haveInitialOrders As Boolean, droidType As Long, i As Long
    Dim initialOrders_ As INITIAL_DROID_ORDERS

    If Not bMultiMessages Then
        SendDroid = True
        Exit Function
    End If

    If Not (x <> 0 And y <> 0) Then ASSERT_FAILURE "x != 0 && y != 0", "src/multibot.cpp:311", "SendDroid", "SendDroid: Invalid droid coordinates": SendDroid = False: Exit Function
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/multibot.cpp:312", "SendDroid", "invalid player %u", player: SendDroid = False: Exit Function

    ' Dont send other droids during campaign setup
    If ingame.localJoiningInProgress Then
        SendDroid = True
        Exit Function
    End If

    ' Only send the droid if we are responsible
    If Not myResponsibility(player) Then
        ' Don't build if we are not responsible
        SendDroid = False
        Exit Function
    End If

    debug_ 327, LOG_SYNC, "SendDroid", "Droid sent with id of %u", id
    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_DEBUG_ADD_DROID)
    pos = Vector3i_(x, y, 0)
    haveInitialOrders = hasInitialOrders
    droidType = apsTemplates(pTemplate).droidType

    NETuint8_t w, player
    NETuint32_t w, id
    NETPosition w, pos
    NETwzstring w, apsTemplates(pTemplate).name_
    NETint32_t w, droidType
    NETuint8_t w, apsTemplates(pTemplate).asParts(COMP_BODY)
    NETuint8_t w, apsTemplates(pTemplate).asParts(COMP_BRAIN)
    NETuint8_t w, apsTemplates(pTemplate).asParts(COMP_PROPULSION)
    NETuint8_t w, apsTemplates(pTemplate).asParts(COMP_REPAIRUNIT)
    NETuint8_t w, apsTemplates(pTemplate).asParts(COMP_ECM)
    NETuint8_t w, apsTemplates(pTemplate).asParts(COMP_SENSOR)
    NETuint8_t w, apsTemplates(pTemplate).asParts(COMP_CONSTRUCT)
    NETint8_t w, apsTemplates(pTemplate).numWeaps
    For i = 0 To apsTemplates(pTemplate).numWeaps - 1
        NETuint32_t w, apsTemplates(pTemplate).asWeaps(i)
    Next
    NETbool w, haveInitialOrders
    If haveInitialOrders Then
        initialOrders_ = initialOrders
        NETuint32_t w, initialOrders_.secondaryOrder
        NETint32_t w, initialOrders_.moveToX
        NETint32_t w, initialOrders_.moveToY
        NETuint32_t w, initialOrders_.factoryId  ' For making scripts happy.
    End If
    debug_ 361, LOG_LIFE, "SendDroid", "===> sending Droid from %u id of %u ", player, id
    SendDroid = NETend(w)
End Function

' ////////////////////////////////////////////////////////////////////////////
' receive droid creation information from other players
Public Function recvDroid(ByVal queue As Long) As Boolean
    Dim pT As Long, psDroid As Long, player As Long, id As Long, pos As Vector3i
    Dim haveInitialOrders As Boolean, initialOrders As INITIAL_DROID_ORDERS, r As MessageReader
    Dim droidType As Long, name_ As String, i As Long, rot As Rotation

    pT = newTemplate()                  ' DROID_TEMPLATE t, *pT = &t (deleted at every return)
    psDroid = 0
    player = 0
    id = 0
    pos = Vector3i_(0, 0, 0)
    haveInitialOrders = False

    Set r = NETbeginDecode(queue, GAME_DEBUG_ADD_DROID)
    NETuint8_t r, player
    NETuint32_t r, id
    NETPosition r, pos
    NETwzstring r, name_
    apsTemplates(pT).name_ = name_
    apsTemplates(pT).id = apsTemplates(pT).name_
    NETint32_t r, droidType
    NETuint8_t r, apsTemplates(pT).asParts(COMP_BODY)
    NETuint8_t r, apsTemplates(pT).asParts(COMP_BRAIN)
    NETuint8_t r, apsTemplates(pT).asParts(COMP_PROPULSION)
    NETuint8_t r, apsTemplates(pT).asParts(COMP_REPAIRUNIT)
    NETuint8_t r, apsTemplates(pT).asParts(COMP_ECM)
    NETuint8_t r, apsTemplates(pT).asParts(COMP_SENSOR)
    NETuint8_t r, apsTemplates(pT).asParts(COMP_CONSTRUCT)
    NETint8_t r, apsTemplates(pT).numWeaps
    If Not (apsTemplates(pT).numWeaps >= 0 And apsTemplates(pT).numWeaps <= MAX_WEAPONS) Then ASSERT_FAILURE "pT->numWeaps >= 0 && pT->numWeaps <= ARRAY_SIZE(pT->asWeaps)", "src/multibot.cpp:397", "recvDroid", "Bad numWeaps %d", apsTemplates(pT).numWeaps: recvDroid = False: GoTo done
    For i = 0 To apsTemplates(pT).numWeaps - 1
        NETuint32_t r, apsTemplates(pT).asWeaps(i)
    Next
    NETbool r, haveInitialOrders
    If haveInitialOrders Then
        NETuint32_t r, initialOrders.secondaryOrder
        NETint32_t r, initialOrders.moveToX
        NETint32_t r, initialOrders.moveToY
        NETuint32_t r, initialOrders.factoryId  ' For making scripts happy.
    End If
    apsTemplates(pT).droidType = droidType
    NETend r

    If Not debugMappingsAllowed() And bMultiPlayer Then
        debug_ 417, LOG_WARNING, "recvDroid", "Failed to add droid for player %u.", NetPlay_players_position(NETQUEUE_index(queue))
        recvDroid = False
        GoTo done
    End If

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/multibot.cpp:421", "recvDroid", "invalid player %u", player: recvDroid = False: GoTo done

    debug_ 423, LOG_LIFE, "recvDroid", "<=== getting Droid from %u id of %u ", player, id
    If (pos.x = 0 And pos.y = 0) Or pos.x > world_coord(mapWidth) Or pos.y > world_coord(mapHeight) Then
        debug_ 426, LOG_ERROR, "recvDroid", "Received bad droid position (%d, %d) from %d about p%d (%s)", pos.x, pos.y, _
              NETQUEUE_index(queue), player, IIf(isHumanPlayer(player), "Human", "AI")
        recvDroid = False
        GoTo done
    End If

    ' Create that droid on this machine.
    psDroid = reallyBuildDroid(pT, pos, player, False, rot, id)

    ' If we were able to build the droid set it up
    If psDroid <> 0 Then
        addDroid psDroid, apsDroidLists

        If haveInitialOrders Then
            aDroid(psDroid).secondaryOrder = initialOrders.secondaryOrder
            aDroid(psDroid).secondaryOrderPending = aDroid(psDroid).secondaryOrder
            orderDroidLoc psDroid, DORDER_MOVE, initialOrders.moveToX, initialOrders.moveToY, ModeImmediate
            cbNewDroid IdToStruct(initialOrders.factoryId, ANYPLAYER), psDroid
        End If

        syncDebugDroid_ "recvDroid", psDroid, Asc("+")
    Else
        debug_ 452, LOG_ERROR, "recvDroid", "Packet from %d cannot create droid for p%d (%s)!", NETQUEUE_index(queue), _
              player, IIf(isHumanPlayer(player), "Human", "AI")
        ' (CONPRINTF: #ifdef DEBUG)
        recvDroid = False
        GoTo done
    End If

    recvDroid = True
done:
    deleteTemplate pT
End Function

' Does not read/write info->droidId!
' (VB) template <typename SerdeContext>: c is the MessageReader or the MessageWriter
Private Sub NETQueuedDroidInfo(ByVal c As Object, info As QueuedDroidInfo)
    NETuint8_t c, info.player
    NETenum c, info.subType
    Select Case info.subType
    Case ObjOrder, LocOrder
        NETenum c, info.order
        If info.subType = ObjOrder Then
            NETuint32_t c, info.destId
            NETenum c, info.destType
        Else
            NETVector2i c, info.pos
        End If
        If info.order = DORDER_BUILD Or info.order = DORDER_LINEBUILD Then
            NETuint32_t c, info.structRef
            NETuint16_t c, info.direction
        End If
        If info.order = DORDER_LINEBUILD Then
            NETVector2i c, info.pos2
        End If
        If info.order = DORDER_BUILDMODULE Then
            NETuint32_t c, info.index
        End If
        NETbool c, info.add
    Case SecondaryOrder
        NETenum c, info.secOrder
        NETenum c, info.secState
    End Select
End Sub

' Actually send the droid info.
Public Sub sendQueuedDroidInfo()
    ' Given an order type, we bring all other orders of the same type together
    ' WHILE KEEPING their relative order!! This is important because
    ' sending all MOVE and then all HOLD_POSITION is clearly not the same as
    ' sending all HOLD_POSITION and then MOVE
    ' yes, we need an ordered map here, not std::sort
    ' (VB) orderedMap: keys() the distinct orders, ascending; qOrders() the indices of one key's infos, in queue order
    Dim keys() As Long, nKeys As Long, k As Long, i As Long, j As Long, key As Long
    Dim qOrders() As Long, nq As Long, eqBegin As Long, eqEnd As Long
    Dim w As MessageWriter, num As Long, prevDroidId As Long, n As Long, droidId As Long, deltaDroidId As Long

    If queuedOrders_size = 0 Then GoTo sent
    ReDim keys(0 To queuedOrders_size - 1)
    ReDim qOrders(0 To queuedOrders_size - 1)
    nKeys = 0
    For i = 0 To queuedOrders_size - 1
        key = queuedOrders(i).order
        k = 0
        Do While k < nKeys
            If keys(k) >= key Then Exit Do
            k = k + 1
        Loop
        If k = nKeys Then
            keys(nKeys) = key
            nKeys = nKeys + 1
        ElseIf keys(k) <> key Then
            For j = nKeys To k + 1 Step -1
                keys(j) = keys(j - 1)
            Next
            keys(k) = key
            nKeys = nKeys + 1
        End If
    Next
    For k = 0 To nKeys - 1
        nq = 0
        For i = 0 To queuedOrders_size - 1
            If queuedOrders(i).order = keys(k) Then
                qOrders(nq) = i
                nq = nq + 1
            End If
        Next
        eqBegin = 0
        Do While eqBegin <> nq
            ' Find end of range of orders which differ only by the droid ID.
            eqEnd = eqBegin + 1
            Do While eqEnd <> nq
                If QueuedDroidInfo_orderCompare(queuedOrders(qOrders(eqEnd)), queuedOrders(qOrders(eqBegin))) <> 0 Then Exit Do
                eqEnd = eqEnd + 1
            Loop

            Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_DROIDINFO)
            NETQueuedDroidInfo w, queuedOrders(qOrders(eqBegin))

            num = eqEnd - eqBegin
            NETuint32_t w, num

            prevDroidId = 0
            For n = 0 To num - 1
                droidId = queuedOrders(qOrders(eqBegin + n)).droidId

                ' Encode deltas between droid IDs, since the deltas are smaller than the actual droid IDs, and will encode to less bytes on average.
                deltaDroidId = FromU32(CDbl(droidId) - prevDroidId)
                NETuint32_t w, deltaDroidId

                prevDroidId = droidId
            Next
            NETend w
            eqBegin = eqEnd
        Loop
    Next
sent:
    ' Sent the orders. Don't send them again.
    queuedOrders_size = 0
End Sub

' (VB) psStats: a STRUCTURE_STATS index, -1 for nullptr
Private Function infoToOrderData(info As QueuedDroidInfo, ByVal psStats As Long) As DROID_ORDER_DATA
    Dim sOrder As DROID_ORDER_DATA
    sOrder = DroidOrder_()
    sOrder.typ = info.order
    sOrder.pos = info.pos
    sOrder.pos2 = info.pos2
    sOrder.direction = info.direction
    sOrder.index = info.index
    sOrder.psObj = processDroidTarget(info.destType, info.destId)
    sOrder.psStats = psStats

    infoToOrderData = sOrder
End Function

' ////////////////////////////////////////////////////////////////////////////
' Droid update information
Public Sub sendDroidInfo(ByVal psDroid As Long, order As DROID_ORDER_DATA, ByVal add As Boolean)
    Dim info As QueuedDroidInfo, sOrder As DROID_ORDER_DATA

    If Not myResponsibility(aDroid(psDroid).player) Then
        Exit Sub
    End If
    If NETisReplay() Then
        Exit Sub
    End If

    info.player = aDroid(psDroid).player
    info.droidId = aDroid(psDroid).id
    If order.psObj <> 0 Then info.subType = ObjOrder Else info.subType = LocOrder
    info.order = order.typ
    If info.subType = ObjOrder Then
        info.destId = SIMPLE_OBJECT_id(order.psObj)
        info.destType = SIMPLE_OBJECT_type(order.psObj)
    Else
        info.pos = order.pos
    End If
    If order.typ = DORDER_BUILD Or order.typ = DORDER_LINEBUILD Then
        info.structRef = asStructureStats(order.psStats).ref
        info.direction = order.direction
        If Not DROID_isConstructionDroid(psDroid) Then
            Exit Sub  ' No point ordering things to build if they can't build anything.
        End If
    End If
    If order.typ = DORDER_LINEBUILD Then
        info.pos2 = order.pos2
    End If
    If order.typ = DORDER_BUILDMODULE Then
        info.index = order.index
    End If

    info.add = add

    ' Send later, grouped by order, so multiple droids with the same order can be encoded to much less data.
    queuedOrders_push_back info

    ' Update pending orders, so the UI knows it happened.
    sOrder = infoToOrderData(info, order.psStats)
    If Not add Then
        aDroid(psDroid).listPendingBegin = aDroid(psDroid).asOrderList_size
    End If
    orderDroidAddPending psDroid, sOrder
End Sub

' ////////////////////////////////////////////////////////////////////////////
' receive droid information form other players.
Public Function recvDroidInfo(ByVal queue As Long) As Boolean
    Dim r As MessageReader, info As QueuedDroidInfo, psStats As Long, typeIndex As Long
    Dim sOrder As DROID_ORDER_DATA, num As Long, n As Long, deltaDroidId As Long, psDroid As Long

    Set r = NETbeginDecode(queue, GAME_DROIDINFO)
    NETQueuedDroidInfo r, info

    psStats = -1
    If info.subType = LocOrder And (info.order = DORDER_BUILD Or info.order = DORDER_LINEBUILD) Then
        ' Find structure target
        For typeIndex = 0 To numStructureStats - 1
            If asStructureStats(typeIndex).ref = info.structRef Then
                psStats = typeIndex
                Exit For
            End If
        Next
    End If

    Select Case info.subType
    Case ObjOrder:       syncDebug "recvDroidInfo", "Order=%s,%d(%d)", getDroidOrderName(info.order), info.destId, info.destType
    Case LocOrder:       syncDebug "recvDroidInfo", "Order=%s,(%d,%d)", getDroidOrderName(info.order), info.pos.x, info.pos.y
    Case SecondaryOrder: syncDebug "recvDroidInfo", "SecondaryOrder=%d,%08X", info.secOrder, info.secState
    End Select

    sOrder = infoToOrderData(info, psStats)

    num = 0
    NETuint32_t r, num

    n = 0
    Do While ULT(n, num)
        ' Get the next droid ID which is being given this order.
        deltaDroidId = 0
        NETuint32_t r, deltaDroidId
        info.droidId = FromU32(CDbl(info.droidId) + deltaDroidId)

        psDroid = IdToDroid(info.droidId, info.player)
        If psDroid = 0 Then
            debug_ 681, LOG_NEVER, "recvDroidInfo", "Packet from %d refers to non-existent droid %u, [%s : p%d]", _
                  NETQUEUE_index(queue), info.droidId, IIf(isHumanPlayer(info.player), "Human", "AI"), info.player
            syncDebug "recvDroidInfo", "Droid %d missing", info.droidId
            GoTo continue_  ' Can't find the droid, so skip this droid.
        End If
        If Not canGiveOrdersFor(NETQUEUE_index(queue), aDroid(psDroid).player) Then
            debug_ 688, LOG_WARNING, "recvDroidInfo", "Droid order (by %d) for wrong player (%d).", NETQUEUE_index(queue), aDroid(psDroid).player
            syncDebug "recvDroidInfo", "Wrong player."
            GoTo continue_
        End If

        checkDroid psDroid, "src/multibot.cpp:693", "recvDroidInfo", max_check_object_recursion

        syncDebugDroid_ "recvDroidInfo", psDroid, Asc("<")

        Select Case info.subType
        Case ObjOrder, LocOrder
            '
            ' If the current order not is a command order and we are not a
            ' commander yet are in the commander group remove us from it.
            '
            If hasCommander(psDroid) And _
               info.order <> DORDER_RTR And _
               info.order <> DORDER_RTR_SPECIFIED Then
                DROID_GROUP_remove aDroid(psDroid).psGroup, psDroid
            End If

            If sOrder.psObj <> TargetMissing Then  ' Only do order if the target didn't die.
                If Not info.add Then
                    orderDroidListEraseRange psDroid, 0, aDroid(psDroid).listSize + 1  ' Clear all non-pending orders, plus the first pending order (which is probably the order we just received).
                    orderDroidBase psDroid, sOrder  ' Execute the order immediately (even if in the middle of another order.
                Else
                    orderDroidAdd psDroid, sOrder   ' Add the order to the (non-pending) list. Will probably overwrite the corresponding pending order, assuming all pending orders were written to the list.
                End If
            End If
        Case SecondaryOrder
            ' Set the droids secondary order
            turnOffMultiMsg True
            secondarySetState psDroid, info.secOrder, info.secState
            turnOffMultiMsg False
        End Select

        syncDebugDroid_ "recvDroidInfo", psDroid, Asc(">")

        checkDroid psDroid, "src/multibot.cpp:735", "recvDroidInfo", max_check_object_recursion
continue_:
        n = n + 1
    Loop
    NETend r

    recvDroidInfo = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' process droid order
Private Function processDroidTarget(ByVal desttype As Long, ByVal destid As Long) As Long
    Dim psObj As Long

    ' Target is a location
    If destid = 0 And desttype = 0 Then
        processDroidTarget = 0
        Exit Function
    ' Target is an object
    Else
        psObj = 0

        Select Case desttype
        Case OBJ_DROID
            psObj = IdToDroid(destid, ANYPLAYER)
        Case OBJ_STRUCTURE
            psObj = IdToStruct(destid, ANYPLAYER)
        Case OBJ_FEATURE
            psObj = IdToFeature(destid, ANYPLAYER)

        ' We should not get this!
        Case OBJ_PROJECTILE
            debug_ 771, LOG_ERROR, "processDroidTarget", "ProcessDroidOrder: order specified destination as a bullet. what am i to do??"
        Case Else
            debug_ 774, LOG_ERROR, "processDroidTarget", "ProcessDroidOrder: unknown object type"
        End Select

        ' If we did not find anything, return
        If psObj = 0 Then                                           ' failed to find it;
            syncDebug "processDroidTarget", "Target missing"
            processDroidTarget = TargetMissing  ' Can't return NULL, since then the order would still be attempted.
            Exit Function
        End If

        processDroidTarget = psObj
    End If
End Function


' ////////////////////////////////////////////////////////////////////////////
' Inform other players that a droid has been destroyed
Public Function SendDestroyDroid(ByVal psDroid As Long) As Boolean
    Dim w As MessageWriter, id As Long

    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_DEBUG_REMOVE_DROID)
    id = aDroid(psDroid).id

    ' Send the droid's ID
    debug_ 799, LOG_DEATH, "SendDestroyDroid", "Requested all players to destroy droid %u", id
    NETuint32_t w, id
    SendDestroyDroid = NETend(w)
End Function

' ////////////////////////////////////////////////////////////////////////////
' Accept a droid which was destroyed on another machine
Public Function recvDestroyDroid(ByVal queue As Long) As Boolean
    Dim psDroid As Long, r As MessageReader, id As Long

    Set r = NETbeginDecode(queue, GAME_DEBUG_REMOVE_DROID)

    ' Retrieve the droid
    NETuint32_t r, id
    psDroid = IdToDroid(id, ANYPLAYER)
    If psDroid = 0 Then
        debug_ 820, LOG_DEATH, "recvDestroyDroid", "droid %d on request from player %d can't be found? Must be dead already?", _
              id, NETQUEUE_index(queue)
        recvDestroyDroid = False
        Exit Function
    End If
    NETend r

    If Not debugMappingsAllowed() And bMultiPlayer Then
        debug_ 830, LOG_WARNING, "recvDestroyDroid", "Failed to remove droid for player %u.", NetPlay_players_position(NETQUEUE_index(queue))
        recvDestroyDroid = False
        Exit Function
    End If

    ' If the droid has not died on our machine yet, destroy it
    If aDroid(psDroid).died = 0 Then
        turnOffMultiMsg True
        debug_ 838, LOG_DEATH, "recvDestroyDroid", "Killing droid %d on request from player %d - huh?", aDroid(psDroid).id, NETQUEUE_index(queue)
        destroyDroid psDroid, FromU32(CDbl(gameTime) - deltaGameTime + 1)  ' deltaGameTime is actually 0 here, since we're between updates. However, the value of gameTime - deltaGameTime + 1 will not change when we start the next tick.
        turnOffMultiMsg False
    Else
        debug_ 844, LOG_DEATH, "recvDestroyDroid", "droid %d is confirmed dead by player %d.", aDroid(psDroid).id, NETQUEUE_index(queue)
    End If

    recvDestroyDroid = True
End Function
