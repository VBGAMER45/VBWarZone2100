Attribute VB_Name = "multigifts_c"
' src/multigifts.cpp, multigifts.h, whole: the alliance code
' (requestAlliance, breakAlliance, formAlliance, sendAlliance, recvAlliance,
' createTeamAlliances) and the gifts: recvGift, sendGift, giftAutoGame,
' giftRadar, recvGiftStruct, recvGiftDroids, sendGiftDroids, giftResearch,
' giftPower, technologyGiveAway, sendMultiPlayerFeature,
' recvMultiPlayerFeature, pickupArtefact.
' recvAlliance's lambda prohibitedNewAlliance is a function here; it logs
' under recvAlliance's name (GCC's __FUNCTION__ in a lambda is
' "operator()"). pickATileGen's function pointer zonedPAT is droid_c's
' PAT_zonedPAT; getLocalizedStatsName and getStatsName of a stats are its
' name_.
Option Explicit

' multigifts.h
Public Const RADAR_GIFT As Long = 1
Public Const DROID_GIFT As Long = 2
Public Const RESEARCH_GIFT As Long = 3
Public Const POWER_GIFT As Long = 4
Public Const STRUCTURE_GIFT As Long = 5
Public Const AUTOGAME_GIFT As Long = 6  ' Notify others that we are now being controlled by the AI

' ////////////////////////////////////////////////////////////////////////////
' gifts..

Public Function recvGift(ByVal queue As Long) As Boolean
    Dim typ As Long, from As Long, to_ As Long, audioTrack As Long, droidID As Long, r As MessageReader

    Set r = NETbeginDecode(queue, GAME_GIFT)
    NETuint8_t r, typ
    NETuint8_t r, from
    NETuint8_t r, to_
    NETuint32_t r, droidID
    NETend r

    If Not canGiveOrdersFor(NETQUEUE_index(queue), from) Then
        debug_ 82, LOG_WARNING, "recvGift", "Gift (%d) from %d, to %d, queue.index %d", typ, from, to_, NETQUEUE_index(queue)
        syncDebug "recvGift", "Wrong player."
        recvGift = False
        Exit Function
    End If

    If to_ >= MAX_PLAYERS Then
        debug_ 89, LOG_WARNING, "recvGift", "Gift (%d) from %d, to %d (invalid recipient player), queue.index %d", typ, from, to_, NETQUEUE_index(queue)
        syncDebug "recvGift", "Invalid recipient player."
        recvGift = False
        Exit Function
    End If

    If bMultiPlayer And (to_ < NetPlay_players_size()) Then
        If NetPlay_players_isSpectator(to_) Then
            debug_ 98, LOG_WARNING, "recvGift", "Can't gift (%d) from %d, to %d (spectator player), queue.index %d", typ, from, to_, NETQUEUE_index(queue)
            syncDebug "recvGift", "Can't gift to spectator."
            recvGift = False
            Exit Function
        End If
    End If

    ' Handle the gift depending on what it is
    Select Case typ
    Case RADAR_GIFT
        audioTrack = ID_SENSOR_DOWNLOAD
        giftRadar from, to_, False
    Case DROID_GIFT
        audioTrack = ID_UNITS_TRANSFER
        recvGiftDroids from, to_, droidID
    Case STRUCTURE_GIFT
        audioTrack = ID_GIFT
        recvGiftStruct from, to_, droidID
    Case RESEARCH_GIFT
        audioTrack = ID_TECHNOLOGY_TRANSFER
        giftResearch from, to_, False
    Case POWER_GIFT
        audioTrack = ID_POWER_TRANSMIT
        giftPower from, to_, droidID, False
    Case AUTOGAME_GIFT
        audioTrack = ID_SOUND_NEXUS_SYNAPTIC_LINK
        giftAutoGame from, to_, False
    Case Else
        debug_ 132, LOG_ERROR, "recvGift", "recvGift: Unknown Gift recvd"
        recvGift = False
        Exit Function
    End Select

    ' If we are on the receiving end play an audio alert.
    If bMultiPlayer And to_ = selectedPlayer Then
        audio_QueueTrack audioTrack
    End If
    recvGift = True
End Function

Public Function sendGift(ByVal typ As Long, ByVal to_ As Long) As Boolean
    Dim audioTrack As Long

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/multigifts.cpp:149", "sendGift", "Must be a player to send a gift (selectedPlayer: %u", selectedPlayer: sendGift = False: Exit Function

    Select Case typ
    Case RADAR_GIFT
        audioTrack = ID_SENSOR_DOWNLOAD
        giftRadar selectedPlayer, to_, True
    Case DROID_GIFT
        audioTrack = ID_UNITS_TRANSFER
        sendGiftDroids selectedPlayer, to_
    Case RESEARCH_GIFT
        audioTrack = ID_TECHNOLOGY_TRANSFER
        giftResearch selectedPlayer, to_, True
    Case POWER_GIFT
        audioTrack = ID_POWER_TRANSMIT
        giftPower selectedPlayer, to_, 0, True
    Case AUTOGAME_GIFT
        giftAutoGame selectedPlayer, to_, True
        sendGift = True
        Exit Function
    Case Else   ' STRUCTURE_GIFT: not implemented
        debug_ 176, LOG_ERROR, "sendGift", "Unknown Gift sent"
        sendGift = False
        Exit Function
    End Select

    ' Play the appropriate audio track
    audio_QueueTrack audioTrack

    sendGift = True
End Function

Private Sub giftAutoGame(ByVal from As Long, ByVal to_ As Long, ByVal send As Boolean)
    Dim dummy As Long, subType As Long, w As MessageWriter, f As Long, t As Long, s As String
    dummy = 0

    If send Then
        subType = AUTOGAME_GIFT

        Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_GIFT)
        f = from
        t = to_
        NETuint8_t w, subType
        NETuint8_t w, f
        NETuint8_t w, t
        NETuint32_t w, dummy
        NETend w
        debug_ 201, LOG_SYNC, "giftAutoGame", "We (%d) are telling %d we want to enable/disable a autogame", from, to_
    ' If we are receiving the "gift"
    Else
        If to_ = selectedPlayer Then
            NetPlay_players_autoGame(from) = Not NetPlay_players_autoGame(from)
            If NetPlay_players_autoGame(from) Then s = "Enabled" Else s = "Disabled"
            CONPRINTF "%s has %s the autoGame command", getPlayerName(from), s
            debug_ 210, LOG_SYNC, "giftAutoGame", "We (%d) are being told that %d has %s autogame", selectedPlayer, from, s
        End If
    End If
End Sub

' ////////////////////////////////////////////////////////////////////////////
' give radar information
Public Sub giftRadar(ByVal from As Long, ByVal to_ As Long, ByVal send As Boolean)
    Dim dummy As Long, subType As Long, w As MessageWriter, f As Long, t As Long
    dummy = 0

    If send Then
        subType = RADAR_GIFT

        Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_GIFT)
        f = from
        t = to_
        NETuint8_t w, subType
        NETuint8_t w, f
        NETuint8_t w, t
        NETuint32_t w, dummy
        NETend w
    ' If we are receiving the gift
    Else
        hqReward from, to_
        If to_ = selectedPlayer And loopMissionState = LMS_NORMAL Then
            CONPRINTF "%s Gives You A Visibility Report", getPlayerName(from)
        End If
    End If
End Sub

' NOTICE: the packet is already set-up for decoding via recvGift()
Private Sub recvGiftStruct(ByVal from As Long, ByVal to_ As Long, ByVal structID As Long)
    Dim psStruct As Long
    psStruct = IdToStruct(structID, from)
    If psStruct <> 0 Then
        syncDebugStructure_ "recvGiftStruct", psStruct, Asc("<")
        giftSingleStructure psStruct, to_, False
        syncDebugStructure_ "recvGiftStruct", psStruct, Asc(">")
        If to_ = selectedPlayer Then
            CONPRINTF "%s Gives you a %s", getPlayerName(from), asStructureStats(aStruct(psStruct).pStructureType).name_
        End If
    Else
        debug_ 260, LOG_ERROR, "recvGiftStruct", "Bad structure id %u, from %u to %u", structID, from, to_
    End If
End Sub

' recvGiftDroid()
' We received a droid gift from another player.
' NOTICE: the packet is already set-up for decoding via recvGift()
'
' \param from  :player that sent us the droid
' \param to    :player that should be getting the droid
Private Sub recvGiftDroids(ByVal from As Long, ByVal to_ As Long, ByVal droidID As Long)
    Dim psDroid As Long, successfulTransfer As Boolean, noPos As Vector2i
    psDroid = IdToDroid(droidID, from)

    If psDroid <> 0 Then
        syncDebugDroid_ "recvGiftDroids", psDroid, Asc("<")
        successfulTransfer = giftSingleDroid(psDroid, to_, False, noPos) <> 0
        syncDebugDroid_ "recvGiftDroids", psDroid, Asc(">")
        If successfulTransfer And (to_ = selectedPlayer) Then
            CONPRINTF "%s Gives you a %s", getPlayerName(from), aDroid(psDroid).aName
        End If
    Else
        debug_ 286, LOG_ERROR, "recvGiftDroids", "Bad droid id %u, from %u to %u", droidID, from, to_
    End If
End Sub

' sendGiftDroids()
' We give selected droid(s) as a gift to another player.
'
' \param from  :player that sent us the droid
' \param to    :player that should be getting the droid
Private Sub sendGiftDroids(ByVal from As Long, ByVal to_ As Long)
    Dim psD As Long, giftType As Long, totalToSend As Long, w As MessageWriter, f As Long, t As Long, id As Long
    giftType = DROID_GIFT
    totalToSend = 0

    If apsDroidLists(from).size = 0 Then
        Exit Sub
    End If

    '
    ' Work out the number of droids to send. As well as making sure they are
    ' selected we also need to make sure they will NOT put the receiving player
    ' over their droid limit.
    '

    psD = apsDroidLists(from).head
    Do While psD <> 0
        If Not (getNumDroids(to_) + totalToSend < getMaxDroids(to_) And totalToSend <> UINT8_MAX) Then Exit Do
        If aDroid(lnObj(psD)).selected <> 0 Then
            ' (&&: transporterIsEmpty, whose ASSERT may log, only for a transporter)
            If DROID_isTransporter(lnObj(psD)) Then
                If Not transporterIsEmpty(lnObj(psD)) Then
                    CONPRINTF "Tried to give away a non-empty %s - but this is not allowed.", aDroid(lnObj(psD)).aName
                    GoTo continue1_
                End If
            End If
            totalToSend = totalToSend + 1
        End If
continue1_:
        psD = lnNext(psD)
    Loop
    '
    ' We must send one droid at a time, due to the fact that giftSingleDroid()
    ' does its own net calls.
    '

    psD = apsDroidLists(from).head
    Do While psD <> 0 And totalToSend <> 0
        If aDroid(lnObj(psD)).selected <> 0 Then
            If DROID_isTransporter(lnObj(psD)) Then
                If Not transporterIsEmpty(lnObj(psD)) Then
                    GoTo continue2_
                End If
            End If

            Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_GIFT)
            f = from
            t = to_
            NETuint8_t w, giftType
            NETuint8_t w, f
            NETuint8_t w, t
            ' Add the droid to the packet
            id = aDroid(lnObj(psD)).id
            NETuint32_t w, id
            NETend w

            ' Decrement the number of droids left to send
            totalToSend = totalToSend - 1
        End If
continue2_:
        psD = lnNext(psD)
    Loop
End Sub


' ////////////////////////////////////////////////////////////////////////////
' give technologies.
Private Sub giftResearch(ByVal from As Long, ByVal to_ As Long, ByVal send As Boolean)
    Dim dummy As Long, giftType As Long, w As MessageWriter, f As Long, t As Long, i As Long
    dummy = 0

    If send Then
        giftType = RESEARCH_GIFT

        Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_GIFT)
        f = from
        t = to_
        NETuint8_t w, giftType
        NETuint8_t w, f
        NETuint8_t w, t
        NETuint32_t w, dummy
        NETend w
    ElseIf alliancesCanGiveResearchAndRadar(game.alliance) Then
        If to_ = selectedPlayer Then
            CONPRINTF "%s Gives You Technology Documents", getPlayerName(from)
        End If
        ' For each topic
        For i = 0 To asResearch_size - 1
            ' If they have it and we don't research it
            If IsResearchCompleted(asPlayerResList(from, i)) Then
                If Not IsResearchCompleted(asPlayerResList(to_, i)) Then
                    researchResult i, to_, False, 0, True
                End If
            End If
        Next
    End If
End Sub


' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' alliance code......

Public Sub giftPower(ByVal from As Long, ByVal to_ As Long, ByVal amount As Long, ByVal send As Boolean)
    Dim giftType As Long, w As MessageWriter, value As Long
    If send Then
        giftType = POWER_GIFT

        Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_GIFT)
        NETuint8_t w, giftType
        NETuint8_t w, from
        NETuint8_t w, to_
        NETuint32_t w, amount
        NETend w
    Else
        value = 0

        If from = ANYPLAYER And Not NetPlay_bComms Then
            ' basically cheating power, so we check that we are not in multiplayer
            addPower to_, amount
        ElseIf from = ANYPLAYER And NetPlay_bComms Then
            debug_ 418, LOG_WARNING, "giftPower", "Someone tried to cheat power in multiplayer - ignored!"
        ElseIf amount = 0 Then ' the GUI option
            value = getPower(from) \ 3
            usePower from, value
            addPower to_, value
        Else ' for scripts etc that can give precise amounts
            ' MIN(getPower(from), amount): int against uint32_t compares unsigned
            If ULT(getPower(from), amount) Then value = getPower(from) Else value = amount
            usePower from, value
            addPower to_, value
        End If
        If from <> ANYPLAYER And to_ = selectedPlayer Then
            CONPRINTF "%s Gives You %d Power", getPlayerName(from), value
        End If
    End If
End Sub


Public Sub requestAlliance(ByVal from As Long, ByVal to_ As Long, ByVal prop As Boolean, ByVal allowAudio As Boolean)
    If prop And bMultiMessages Then
        sendAlliance from, to_, ALLIANCE_REQUESTED, 0
        Exit Sub  ' Wait for our message.
    End If

    syncDebug "requestAlliance", "Request alliance %d %d", from, to_
    alliances(from, to_) = ALLIANCE_REQUESTED  ' We've asked
    alliances(to_, from) = ALLIANCE_INVITATION ' They've been invited


    triggerEventAllianceOffer from, to_

    If to_ = selectedPlayer Then
        CONPRINTF "%s Requests An Alliance With You", getPlayerName(from)

        If allowAudio Then
            audio_QueueTrack ID_ALLIANCE_OFF
        End If
    ElseIf from = selectedPlayer Then
        CONPRINTF "You Invite %s To Form An Alliance", getPlayerName(to_)
        If allowAudio Then
            audio_QueueTrack ID_ALLIANCE_OFF
        End If
    End If
End Sub

Public Sub breakAlliance(ByVal p1 As Long, ByVal p2 As Long, ByVal prop As Boolean, ByVal allowAudio As Boolean)
    Dim tm1 As String, it As Long, psStructure As Long, b As StructureBounds, i As Long, j As Long

    If prop And bMultiMessages Then
        sendAlliance p1, p2, ALLIANCE_BROKEN, 0
        Exit Sub  ' Wait for our message.
    End If

    If alliances(p1, p2) = ALLIANCE_FORMED Then
        tm1 = getPlayerName(p1)
        If Len(tm1) > 127 Then tm1 = Left$(tm1, 127)
        CONPRINTF "%s Breaks The Alliance With %s", tm1, getPlayerName(p2)

        If allowAudio And (p1 = selectedPlayer Or p2 = selectedPlayer) Then
            audio_QueueTrack ID_ALLIANCE_BRO
        End If
    End If

    syncDebug "breakAlliance", "Break alliance %d %d", p1, p2
    triggerEventAllianceBroken p1, p2
    alliances(p1, p2) = ALLIANCE_BROKEN
    alliances(p2, p1) = ALLIANCE_BROKEN
    alliancebits(p1) = alliancebits(p1) And Not BIT_(p2)
    alliancebits(p2) = alliancebits(p2) And Not BIT_(p1)

    ' Make sure p1's structures are no longer considered "our buildings" to their former allies
    ' For unit pathing
    it = apsStructLists(p1).head
    Do While it <> 0
        psStructure = lnObj(it)
        b = getStructureBounds_struct(psStructure)

        For i = 0 To b.size.x - 1
            For j = 0 To b.size.y - 1
                auxClearAll b.map.x + i, b.map.y + j, AUXBITS_OUR_BUILDING
                auxSetAllied b.map.x + i, b.map.y + j, aStruct(psStructure).player, AUXBITS_OUR_BUILDING
            Next
        Next
        it = lnNext(it)
    Loop
    ' Do the same for p2's stuff
    it = apsStructLists(p2).head
    Do While it <> 0
        psStructure = lnObj(it)
        b = getStructureBounds_struct(psStructure)

        For i = 0 To b.size.x - 1
            For j = 0 To b.size.y - 1
                auxClearAll b.map.x + i, b.map.y + j, AUXBITS_OUR_BUILDING
                auxSetAllied b.map.x + i, b.map.y + j, aStruct(psStructure).player, AUXBITS_OUR_BUILDING
            Next
        Next
        it = lnNext(it)
    Loop
End Sub

Public Sub formAlliance(ByVal p1 As Long, ByVal p2 As Long, ByVal prop As Boolean, ByVal allowAudio As Boolean, ByVal allowNotification As Boolean)
    Dim tm1 As String, it As Long, psDroid As Long, psStructure As Long, b As StructureBounds, i As Long, j As Long

    If bMultiMessages And prop Then
        sendAlliance p1, p2, ALLIANCE_FORMED, 0
        Exit Sub  ' Wait for our message.
    End If

    ' Don't add message if already allied
    If bMultiPlayer And alliances(p1, p2) <> ALLIANCE_FORMED And allowNotification Then
        tm1 = getPlayerName(p1)
        If Len(tm1) > 127 Then tm1 = Left$(tm1, 127)
        CONPRINTF "%s Forms An Alliance With %s", tm1, getPlayerName(p2)
    End If

    syncDebug "formAlliance", "Form alliance %d %d", p1, p2
    triggerEventAllianceAccepted p1, p2
    alliances(p1, p2) = ALLIANCE_FORMED
    alliances(p2, p1) = ALLIANCE_FORMED
    If bMultiPlayer And alliancesSharedVision(game.alliance) Then  ' this is for shared vision only
        alliancebits(p1) = alliancebits(p1) Or BIT_(p2)
        alliancebits(p2) = alliancebits(p2) Or BIT_(p1)
    End If

    If allowAudio And (p1 = selectedPlayer Or p2 = selectedPlayer) Then
        audio_QueueTrack ID_ALLIANCE_ACC
    End If

    ' Not campaign and alliances are transitive
    If bMultiPlayer And alliancesSharedVision(game.alliance) Then
        giftRadar p1, p2, False
        giftRadar p2, p1, False
    End If

    ' Clear out any attacking orders
    it = apsDroidLists(p1).head                 ' from -> to
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).order.typ = DORDER_ATTACK Then
            If aDroid(psDroid).order.psObj <> 0 Then
                If SIMPLE_OBJECT_player(aDroid(psDroid).order.psObj) = p2 Then
                    orderDroid psDroid, DORDER_STOP, ModeImmediate
                End If
            End If
        End If
        it = lnNext(it)
    Loop
    it = apsDroidLists(p2).head                 ' to -> from
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).order.typ = DORDER_ATTACK Then
            If aDroid(psDroid).order.psObj <> 0 Then
                If SIMPLE_OBJECT_player(aDroid(psDroid).order.psObj) = p1 Then
                    orderDroid psDroid, DORDER_STOP, ModeImmediate
                End If
            End If
        End If
        it = lnNext(it)
    Loop

    ' Properly mark all of p1's structures as allied buildings for unit pathing
    it = apsStructLists(p1).head
    Do While it <> 0
        psStructure = lnObj(it)
        b = getStructureBounds_struct(psStructure)

        For i = 0 To b.size.x - 1
            For j = 0 To b.size.y - 1
                If Not (asStructureStats(aStruct(psStructure).pStructureType).typ = REF_GATE) Then
                    auxSetAllied b.map.x + i, b.map.y + j, aStruct(psStructure).player, AUXBITS_OUR_BUILDING
                Else
                    ' Make sure gates aren't set as impassible to our new allies
                    auxClearAll b.map.x + i, b.map.y + j, AUXBITS_NONPASSABLE
                    auxSetEnemy b.map.x + i, b.map.y + j, aStruct(psStructure).player, AUXBITS_NONPASSABLE
                End If
            Next
        Next
        it = lnNext(it)
    Loop
    ' Do the same for p2's stuff
    it = apsStructLists(p2).head
    Do While it <> 0
        psStructure = lnObj(it)
        b = getStructureBounds_struct(psStructure)

        For i = 0 To b.size.x - 1
            For j = 0 To b.size.y - 1
                If Not (asStructureStats(aStruct(psStructure).pStructureType).typ = REF_GATE) Then
                    auxSetAllied b.map.x + i, b.map.y + j, aStruct(psStructure).player, AUXBITS_OUR_BUILDING
                Else
                    auxClearAll b.map.x + i, b.map.y + j, AUXBITS_NONPASSABLE
                    auxSetEnemy b.map.x + i, b.map.y + j, aStruct(psStructure).player, AUXBITS_NONPASSABLE
                End If
            Next
        Next
        it = lnNext(it)
    Loop
End Sub

Public Sub sendAlliance(ByVal from As Long, ByVal to_ As Long, ByVal state As Long, ByVal value As Long)
    Dim w As MessageWriter
    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_ALLIANCE)
    NETuint8_t w, from
    NETuint8_t w, to_
    NETuint8_t w, state
    NETint32_t w, value
    NETend w
End Sub

' recvAlliance's lambda prohibitedNewAlliance
Private Function prohibitedNewAlliance(ByVal from As Long, ByVal to_ As Long) As Boolean
    If bMultiPlayer Then
        If from < NetPlay_players_size() Then
            If NetPlay_players_isSpectator(from) Then
                debug_ 680, LOG_WARNING, "recvAlliance", "Can't enable alliance from %d (spectator), to %d", from, to_
                syncDebug "recvAlliance", "Can't enable alliance from spectator."
                prohibitedNewAlliance = True
                Exit Function
            End If
        End If
        If to_ < NetPlay_players_size() Then
            If NetPlay_players_isSpectator(to_) Then
                debug_ 686, LOG_WARNING, "recvAlliance", "Can't enable alliance from %d, to %d (spectator)", from, to_
                syncDebug "recvAlliance", "Can't enable alliance to spectator."
                prohibitedNewAlliance = True
                Exit Function
            End If
        End If
    End If
    prohibitedNewAlliance = False
End Function

Public Function recvAlliance(ByVal queue As Long, ByVal allowAudio As Boolean) As Boolean
    Dim to_ As Long, from As Long, state As Long, value As Long, r As MessageReader

    Set r = NETbeginDecode(queue, GAME_ALLIANCE)
    NETuint8_t r, from
    NETuint8_t r, to_
    NETuint8_t r, state
    NETint32_t r, value
    NETend r

    If Not canGiveOrdersFor(NETQUEUE_index(queue), from) Then
        recvAlliance = False
        Exit Function
    End If

    If to_ >= MAX_PLAYERS Then
        debug_ 671, LOG_WARNING, "recvAlliance", "Invalid recipient player (%d), queue.index %d", to_, NETQUEUE_index(queue)
        recvAlliance = False
        Exit Function
    End If

    Select Case state
    Case ALLIANCE_NULL
    Case ALLIANCE_REQUESTED
        If prohibitedNewAlliance(from, to_) Then recvAlliance = False: Exit Function
        requestAlliance from, to_, False, allowAudio
    Case ALLIANCE_FORMED
        If prohibitedNewAlliance(from, to_) Then recvAlliance = False: Exit Function
        formAlliance from, to_, False, allowAudio, True
    Case ALLIANCE_BROKEN
        breakAlliance from, to_, False, allowAudio
    Case Else
        debug_ 710, LOG_ERROR, "recvAlliance", "Unknown alliance state recvd."
        recvAlliance = False
        Exit Function
    End Select

    recvAlliance = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' add an artifact on destruction if required.
Public Sub technologyGiveAway(ByVal pS As Long)
    Dim featureIndex As Long, x As Long, y As Long, pF As Long

    ' If a fully built factory (or with modules under construction) which is our responsibility got destroyed
    If asStructureStats(aStruct(pS).pStructureType).typ = REF_FACTORY And (aStruct(pS).status = SS_BUILT Or Not ULT(aStruct(pS).currentBuildPts, aStruct(pS).body)) Then
        syncDebug "technologyGiveAway", "Adding artefact."
    Else
        syncDebug "technologyGiveAway", "Not adding artefact."
        Exit Sub
    End If

    featureIndex = 0
    Do While featureIndex < asFeatureStats_size
        If asFeatureStats(featureIndex).subType = FEAT_GEN_ARTE Then Exit Do
        featureIndex = featureIndex + 1
    Loop
    If featureIndex >= asFeatureStats_size Then
        debug_ 738, LOG_WARNING, "technologyGiveAway", "No artefact feature!"
        Exit Sub
    End If

    x = map_coord(aStruct(pS).pos.x)
    y = map_coord(aStruct(pS).pos.y)
    If Not pickATileGen(x, y, LOOK_FOR_EMPTY_TILE, PAT_zonedPAT) Then
        syncDebug "technologyGiveAway", "Did not find location for oil drum."
        debug_ 746, LOG_FEATURE, "technologyGiveAway", "Unable to find a free location."
        Exit Sub
    End If
    pF = buildFeature(featureIndex, world_coord(x), world_coord(y), False)
    If pF <> 0 Then
        aFeature(pF).player = aStruct(pS).player
        syncDebugFeature_ "technologyGiveAway", pF, Asc("+")
    Else
        debug_ 757, LOG_ERROR, "technologyGiveAway", "Couldn't build artefact?"
    End If
End Sub

' Sends a build order for the given feature type to all players
'  \param subType the type of feature to build
'  \param x,y the coordinates to place the feature at
Public Sub sendMultiPlayerFeature(ByVal ref As Long, ByVal x As Long, ByVal y As Long, ByVal id As Long)
    Dim w As MessageWriter
    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_DEBUG_ADD_FEATURE)
    NETuint32_t w, ref
    NETuint32_t w, x
    NETuint32_t w, y
    NETuint32_t w, id
    NETend w
End Sub

Public Sub recvMultiPlayerFeature(ByVal queue As Long)
    Dim ref As Long, x As Long, y As Long, id As Long, r As MessageReader, i As Long, end_ As Long
    ref = &HFF&
    x = 0
    y = 0
    id = 0

    Set r = NETbeginDecode(queue, GAME_DEBUG_ADD_FEATURE)
    NETuint32_t r, ref
    NETuint32_t r, x
    NETuint32_t r, y
    NETuint32_t r, id
    NETend r

    If Not debugMappingsAllowed() And bMultiPlayer Then
        debug_ 793, LOG_WARNING, "recvMultiPlayerFeature", "Failed to add feature for player %u.", NetPlay_players_position(NETQUEUE_index(queue))
        Exit Sub
    End If

    ' Find the feature stats list that contains the feature type we want to build
    end_ = asFeatureStats_size
    For i = 0 To end_ - 1
        ' If we found the correct feature type
        If asFeatureStats(i).ref = ref Then
            ' Create a feature of the specified type at the given location
            buildFeature i, x, y, False, id
            Exit For
        End If
    Next
End Sub

' ///////////////////////////////////////////////////////////////
Public Function pickupArtefact(ByVal toPlayer As Long, ByVal fromPlayer As Long) As Boolean
    Dim topic As Long
    If fromPlayer < MAX_PLAYERS And bMultiPlayer Then
        If toPlayer = selectedPlayer And toPlayer = fromPlayer Then
            audio_QueueTrack ID_SOUND_ARTIFACT_RECOVERED
            pickupArtefact = True
            Exit Function
        End If

        For topic = asResearch_size - 1 To 0 Step -1
            If IsResearchCompleted(asPlayerResList(fromPlayer, topic)) Then
                If Not IsResearchPossible(asPlayerResList(toPlayer, topic)) Then
                    ' Make sure the topic can be researched
                    If asResearch(topic).researchPower <> 0 And asResearch(topic).researchPoints <> 0 Then
                        MakeResearchPossible asPlayerResList(toPlayer, topic)
                        If toPlayer = selectedPlayer Then
                            CONPRINTF "You Discover Blueprints For %s", asResearch(topic).name_
                        End If
                        Exit For
                    ' Invalid topic
                    Else
                        debug_ 839, LOG_WARNING, "pickupArtefact", "%s is a invalid research topic?", asResearch(topic).name_
                    End If
                End If
            End If
        Next

        If toPlayer = selectedPlayer Then
            audio_QueueTrack ID_SOUND_ARTIFACT_RECOVERED
        End If

        pickupArtefact = True
        Exit Function
    End If

    pickupArtefact = False
End Function

' Ally team members with each other
Public Sub createTeamAlliances()
    Dim i As Long, j As Long
    debug_ 858, LOG_WZ, "createTeamAlliances", "Creating teams"

    For i = 0 To MAX_PLAYERS - 1
        For j = 0 To MAX_PLAYERS - 1
            ' two different players, ...not scavenger player
            If i <> j _
                And i <> scavengerSlot() _
                And j <> scavengerSlot() Then
                ' ...belonging to the same team, ...not allied and not ignoring teams, ...not disabled, ...not spectators
                If NetPlay_players_team(i) = NetPlay_players_team(j) _
                    And Not aiCheckAlliances(i, j) _
                    And NetPlay_players_difficulty(i) <> AIDifficulty_DISABLED _
                    And NetPlay_players_difficulty(j) <> AIDifficulty_DISABLED _
                    And Not NetPlay_players_isSpectator(i) _
                    And Not NetPlay_players_isSpectator(j) Then
                    ' Create silently
                    formAlliance i, j, False, False, False
                End If
            End If
        Next
    Next
End Sub
