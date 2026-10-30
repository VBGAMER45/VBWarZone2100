Attribute VB_Name = "formation_c"
' src/formation.cpp, formation.h, formationdef.h: Control units moving in
' formation.
' (VB) A FORMATION * is a slot of aFormation (1-based; 0 = nullptr); new
' FORMATION() takes the lowest free slot, zeroed as value-initialisation
' does, and delete frees it. psFormationLists (PerPlayerObjectLists, a
' std::list per player) is a singly linked list through fmNext, from
' fmHead(player), in the std::list's order. F_MEMBER's line and next are
' line_ and next_.
Option Explicit

' maximum number of lines in a formation
Public Const F_MAXLINES As Long = 4
' maximum number of unit members of a formation (cannot be more that 128)
Public Const F_MAXMEMBERS As Long = 20

' information about a formation line
' a linked list of the formation members on this line is maintained
' using their index in the asMembers array.  (-1 == 'NULL')
' (cuts down the memory use over proper pointers)
Public Type F_LINE
    xoffset As Long                     ' SWORD: position relative to center
    yoffset As Long                     ' SWORD
    direction As Long                   ' uint16_t: orientation of line
    member As Long                      ' SBYTE: first member in the 'linked list' of members
End Type

' information about a formation member
Public Type F_MEMBER
    line_ As Long                       ' SBYTE: which line this member is on
    next_ As Long                       ' SBYTE: the next member on this line
    dist As Long                        ' SWORD: distance along the line
    psDroid As Long                     ' DROID *: the member unit
End Type

' information about a formation
Public Type FORMATION
    refCount As Long                    ' SWORD: number of units using the formation

    size As Long                        ' SWORD: maximum length of the lines
    rankDist As Long                    ' SWORD: seperation between the ranks
    direction As Long                   ' uint16_t: direction of the formation
    x As Long                           ' position of the front of the formation
    y As Long

    ' the lines that make up a formation
    asLines(0 To F_MAXLINES - 1) As F_LINE
    numLines As Long                    ' SWORD
    maxRank As Long                     ' UBYTE

    ' the units that have a position allocated in the formation
    free As Long                        ' SBYTE
    asMembers(0 To F_MAXMEMBERS - 1) As F_MEMBER

    ' formation speed (currently speed of slowest member) - GJ - sorry.
    iSpeed As Long                      ' UDWORD

    player As Long                      ' uint32_t
End Type

Public Enum FORMATION_TYPE
    FT_LINE
    FT_COLUMN
End Enum

' radius for the different body sizes
Private Const fmLtRad As Long = 80
Private Const fmMedRad As Long = 100
Private Const fmHvyRad As Long = 110

' default length of a formation line
Private Const F_DEFLENGTH As Long = 4 * fmLtRad

' how close to the formation a unit has to be to receive a formation point
Private Const F_JOINRANGE As Long = TILE_UNITS * 3

' how far apart to keep the ranks in a formation
Private Const RANK_DIST As Long = 2 * fmLtRad

' maximum number of ranks for a line
Private Const MAX_RANK As Long = 15

' the distance for finding formations
Private Const FIND_RANGE As Long = TILE_UNITS \ 2

Private Const FORMATION_SPEED_INIT As Long = 100000

' (VB) the formations
Public aFormation() As FORMATION
Private aFormation_size As Long
Private fmUsed() As Boolean
Private fmNext() As Long

' The list of allocated formations
Private fmHead(0 To MAX_PLAYERS - 1) As Long

' (VB) new FORMATION()
Private Function newFormation() As Long
    Dim f As Long, blank As FORMATION
    For f = 1 To aFormation_size
        If Not fmUsed(f) Then Exit For
    Next
    If f > aFormation_size Then
        aFormation_size = aFormation_size + 16
        ReDim Preserve aFormation(1 To aFormation_size)
        ReDim Preserve fmUsed(1 To aFormation_size)
        ReDim Preserve fmNext(1 To aFormation_size)
    End If
    aFormation(f) = blank
    fmUsed(f) = True
    fmNext(f) = 0
    newFormation = f
End Function

' (VB) delete psFormation
Private Sub deleteFormation(ByVal psFormation As Long)
    fmUsed(psFormation) = False
End Sub

' (VB) psFormationLists[player].remove(psFormation)
Private Sub listRemove(ByVal player As Long, ByVal psFormation As Long)
    Dim it As Long, prev As Long, nxt As Long
    prev = 0
    it = fmHead(player)
    Do While it <> 0
        nxt = fmNext(it)
        If it = psFormation Then
            If prev = 0 Then fmHead(player) = nxt Else fmNext(prev) = nxt
        Else
            prev = it
        End If
        it = nxt
    Loop
End Sub

' Initialise the formation system
Public Function formationInitialise() As Boolean
    Dim player As Long
    For player = 0 To MAX_PLAYERS - 1
        fmHead(player) = 0
    Next

    formationInitialise = True
End Function

' Shutdown the formation system
Public Sub formationShutDown()
    Dim player As Long, psCurr As Long
    For player = 0 To MAX_PLAYERS - 1
        psCurr = fmHead(player)
        Do While psCurr <> 0
            debug_ 87, LOG_NEVER, "formationShutDown", "formation with %d units still attached", aFormation(psCurr).refCount
            deleteFormation psCurr
            psCurr = fmNext(psCurr)
        Loop
        fmHead(player) = 0
    Next
End Sub

' Create a new formation
Public Function formationNew(ppsFormation As Long, ByVal player As Long, ByVal typ As FORMATION_TYPE, _
                             ByVal x As Long, ByVal y As Long, ByVal dir_ As Long) As Boolean
    Dim i As Long, psNew As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/formation.cpp:99", "formationNew", "Invalid player: %u", player: formationNew = False: Exit Function

    psNew = newFormation()

    ' get a heap structure
    If Not (psNew <> 0) Then ASSERT_FAILURE "psNew", "src/formation.cpp:105", "formationNew", "Out of memory": formationNew = False: Exit Function

    ' initialise it
    aFormation(psNew).refCount = 0
    aFormation(psNew).size = F_DEFLENGTH
    aFormation(psNew).rankDist = RANK_DIST
    aFormation(psNew).direction = dir_ And &HFFFF&
    aFormation(psNew).x = x
    aFormation(psNew).y = y
    aFormation(psNew).free = 0
    aFormation(psNew).iSpeed = FORMATION_SPEED_INIT
    aFormation(psNew).player = player
    For i = 0 To F_MAXMEMBERS - 1           ' memset(psNew->asMembers, 0, ...)
        aFormation(psNew).asMembers(i).line_ = 0
        aFormation(psNew).asMembers(i).next_ = 0
        aFormation(psNew).asMembers(i).dist = 0
        aFormation(psNew).asMembers(i).psDroid = 0
    Next
    For i = 0 To F_MAXMEMBERS - 1
        aFormation(psNew).asMembers(i).next_ = i + 1
    Next
    aFormation(psNew).asMembers(F_MAXMEMBERS - 1).next_ = -1

    ' add the formation lines based on the formation type
    Select Case typ
    Case FT_LINE
        aFormation(psNew).numLines = 2
        ' line to the left
        aFormation(psNew).asLines(0).xoffset = 0
        aFormation(psNew).asLines(0).yoffset = 0
        aFormation(psNew).asLines(0).direction = (dir_ - DEG(110)) And &HFFFF&
        aFormation(psNew).asLines(0).member = -1
        ' line to the right
        aFormation(psNew).asLines(1).xoffset = 0
        aFormation(psNew).asLines(1).yoffset = 0
        aFormation(psNew).asLines(1).direction = (dir_ + DEG(110)) And &HFFFF&
        aFormation(psNew).asLines(1).member = -1
    Case FT_COLUMN
        aFormation(psNew).numLines = 1
        ' line to the left
        aFormation(psNew).asLines(0).xoffset = 0
        aFormation(psNew).asLines(0).yoffset = 0
        aFormation(psNew).asLines(0).direction = (dir_ + DEG(180)) And &HFFFF&
        aFormation(psNew).asLines(0).member = -1
    Case Else
        ASSERT_FAILURE "false", "src/formation.cpp:149", "formationNew", "Unknown formation type"
        formationNew = False                ' (the C++ leaks psNew)
        Exit Function
    End Select

    fmNext(psNew) = fmHead(player)          ' psFormationLists[player].push_front(psNew)
    fmHead(player) = psNew

    ppsFormation = psNew

    formationNew = True
End Function

' Try to find a formation near a location
Public Function formationFind(ByVal player As Long, ByVal x As Long, ByVal y As Long) As Long
    Dim psFormation As Long, xdiff As Long, ydiff As Long, distSq As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/formation.cpp:165", "formationFind", "Invalid player: %u", player: formationFind = 0: Exit Function

    psFormation = fmHead(player)
    Do While psFormation <> 0
        ' see if the position is close to the formation
        xdiff = aFormation(psFormation).x - x
        ydiff = aFormation(psFormation).y - y
        distSq = xdiff * xdiff + ydiff * ydiff
        If distSq < FIND_RANGE * FIND_RANGE Then
            formationFind = psFormation
            Exit Function
        End If
        psFormation = fmNext(psFormation)
    Loop

    formationFind = 0
End Function

' Find formation speed.
' This currently means the speed of the slowest unit.
Private Sub formationUpdateSpeed(ByVal psFormation As Long, ByVal psNew As Long)
    Dim iUnit As Long, psMember As Long

    If psNew <> 0 Then
        If Not (aDroid(psNew).typ = OBJ_DROID) Then ASSERT_FAILURE "psNew->type == OBJ_DROID", "src/formation.cpp:192", "formationUpdateSpeed", "We've been passed a DROID that really isn't a DROID"
        If ULT(aDroid(psNew).baseSpeed, aFormation(psFormation).iSpeed) Then
            aFormation(psFormation).iSpeed = aDroid(psNew).baseSpeed
        End If
    Else
        aFormation(psFormation).iSpeed = FORMATION_SPEED_INIT
    End If

    For iUnit = 0 To F_MAXMEMBERS - 1
        psMember = aFormation(psFormation).asMembers(iUnit).psDroid
        If psMember <> 0 Then
            If ULT(aDroid(psMember).baseSpeed, aFormation(psFormation).iSpeed) Then
                aFormation(psFormation).iSpeed = aDroid(psMember).baseSpeed
            End If
        End If
    Next
End Sub

' Associate a unit with a formation.
Public Sub formationJoin(ByVal psFormation As Long, ByVal psDroid As Long)
    Dim rankDist As Long, size As Long

    If Not (psFormation <> 0) Then ASSERT_FAILURE "psFormation != nullptr", "src/formation.cpp:219", "formationJoin", "Invalid formation": Exit Sub
    If Not (aFormation(psFormation).player = aDroid(psDroid).player) Then ASSERT_FAILURE "psFormation->player == psDroid->player", "src/formation.cpp:220", "formationJoin", "Formation player (%d) does not match droid player (%d)", aFormation(psFormation).player, aDroid(psDroid).player: Exit Sub

    aFormation(psFormation).refCount = ToInt16(aFormation(psFormation).refCount + 1)

    rankDist = formationObjRadius(psDroid) * 2
    If aFormation(psFormation).rankDist < rankDist Then
        aFormation(psFormation).rankDist = ToInt16(rankDist)
    End If

    size = formationObjRadius(psDroid) * 4
    If aFormation(psFormation).size < size Then
        aFormation(psFormation).size = ToInt16(size)
    End If

    ' update formation speed
    formationUpdateSpeed psFormation, psDroid
End Sub

' Remove a unit from a formation.
Public Sub formationLeave(ByVal psFormation As Long, ByVal psDroid As Long)
    Dim prev As Long, curr As Long, unit As Long, line_ As Long

    If Not (psFormation <> 0) Then ASSERT_FAILURE "psFormation != nullptr", "src/formation.cpp:248", "formationLeave", "Invalid formation": Exit Sub
    If psDroid = 0 Then
        Exit Sub
    End If
    If Not (aFormation(psFormation).refCount > 0) Then ASSERT_FAILURE "psFormation->refCount > 0", "src/formation.cpp:253", "formationLeave", "Formation refcount is zero": Exit Sub

    ' see if the unit is a member
    For unit = 0 To F_MAXMEMBERS - 1
        If aFormation(psFormation).asMembers(unit).psDroid = psDroid Then
            Exit For
        End If
    Next

    If unit < F_MAXMEMBERS Then
        ' remove the member from the members list
        prev = -1
        line_ = aFormation(psFormation).asMembers(unit).line_
        curr = aFormation(psFormation).asLines(line_).member
        Do While curr <> unit
            prev = curr
            curr = aFormation(psFormation).asMembers(curr).next_
        Loop

        If prev = -1 Then
            aFormation(psFormation).asLines(line_).member = aFormation(psFormation).asMembers(unit).next_
        Else
            aFormation(psFormation).asMembers(prev).next_ = aFormation(psFormation).asMembers(unit).next_
        End If
        aFormation(psFormation).asMembers(unit).next_ = aFormation(psFormation).free
        aFormation(psFormation).asMembers(unit).psDroid = 0
        aFormation(psFormation).free = unit

        ' update formation speed
        formationUpdateSpeed psFormation, 0
    End If

    aFormation(psFormation).refCount = ToInt16(aFormation(psFormation).refCount - 1)
    If aFormation(psFormation).refCount = 0 Then
        If ULT(aFormation(psFormation).player, MAX_PLAYERS) Then
            listRemove aFormation(psFormation).player, psFormation
        Else
            If Not (ULT(aFormation(psFormation).player, MAX_PLAYERS)) Then ASSERT_FAILURE "psFormation->player < 11", "src/formation.cpp:304", "formationLeave", "Formation->player is invalid?: %d", aFormation(psFormation).player
        End If
        deleteFormation psFormation
    End If
End Sub

' Remove all the members from a formation and release it.
Public Sub formationReset(ByVal psFormation As Long)
    Dim i As Long

    For i = 0 To F_MAXMEMBERS - 1
        If aFormation(psFormation).asMembers(i).psDroid <> 0 Then
            formationLeave psFormation, aFormation(psFormation).asMembers(i).psDroid
        End If
    Next
End Sub

' Calculate the coordinates of a position on a line
Private Sub formationCalcPos(ByVal psDroid As Long, ByVal psFormation As Long, ByVal line_ As Long, ByVal dist As Long, pX As Long, pY As Long)
    Dim rank As Long, dir_ As Long, xoffset As Long, yoffset As Long
    rank = dist \ aFormation(psFormation).size

    ' calculate the offset of the line based on the rank
    dir_ = (aFormation(psFormation).direction + DEG(180)) And &HFFFF&
    xoffset = iSinR(dir_, aFormation(psFormation).rankDist * rank) + aFormation(psFormation).asLines(line_).xoffset
    yoffset = iCosR(dir_, aFormation(psFormation).rankDist * rank) + aFormation(psFormation).asLines(line_).yoffset

    ' calculate the position of the gap
    dir_ = aFormation(psFormation).asLines(line_).direction
    dist = dist - aFormation(psFormation).size * rank
    pX = iSinR(dir_, dist) + xoffset + aFormation(psFormation).x
    pY = iCosR(dir_, dist) + yoffset + aFormation(psFormation).y
End Sub

' assign a unit to a free spot in the formation
Private Sub formationFindFree(ByVal psFormation As Long, ByVal psDroid As Long, pX As Long, pY As Long)
    Dim line_ As Long, unit As Long, objRadius As Long, radius As Long
    Dim currDist As Long, prev As Long, rank As Long
    Dim x As Long, y As Long, xdiff As Long, ydiff As Long, dist As Long, objDist As Long
    Dim chosenLine As Long, chosenDist As Long, chosenPrev As Long, chosenRank As Long
    Dim found_ As Boolean

    If Not (aFormation(psFormation).free <> -1) Then ASSERT_FAILURE "psFormation->free != -1", "src/formation.cpp:359", "formationFindFree", "No members left to allocate"
    If aFormation(psFormation).free = -1 Then
        pX = aFormation(psFormation).x
        pY = aFormation(psFormation).y
        Exit Sub
    End If

    objRadius = formationObjRadius(psDroid)

    pX = aFormation(psFormation).x
    pY = aFormation(psFormation).y
    chosenLine = 0
    chosenDist = 0
    chosenPrev = -1
    objDist = SDWORD_MAX
    chosenRank = SDWORD_MAX
    For line_ = 0 To aFormation(psFormation).numLines - 1
        ' find the first free position on this line
        currDist = 0
        rank = 1
        prev = -1
        unit = aFormation(psFormation).asLines(line_).member
        found_ = False
        Do While Not found_ And rank <= MAX_RANK
            If Not (unit < F_MAXMEMBERS) Then ASSERT_FAILURE "unit < 20", "src/formation.cpp:386", "formationFindFree", "formationFindFree: unit out of range"

            If unit <> -1 Then
                ' See if the DROID can fit in the gap between this and the last unit
                radius = formationObjRadius(aFormation(psFormation).asMembers(unit).psDroid)
                If objRadius * 2 <= aFormation(psFormation).asMembers(unit).dist - radius - currDist Then
                    found_ = True
                Else
                    prev = unit
                    currDist = aFormation(psFormation).asMembers(unit).dist + radius
                    unit = aFormation(psFormation).asMembers(unit).next_

                    ' if this line is full move onto a rank behind it
                    ' reset the distance to the start of a rank
                    If currDist + objRadius >= aFormation(psFormation).size * rank Then
                        currDist = aFormation(psFormation).size * rank
                        rank = rank + 1
                    End If
                End If
            Else
                found_ = True
            End If

            If found_ Then
                ' calculate the position on the line
                formationCalcPos psDroid, psFormation, line_, currDist + objRadius, x, y
                If fpathBlockingTile(map_coord(x), map_coord(y), asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType) Then
                    ' blocked, try the next rank
                    found_ = False
                    currDist = aFormation(psFormation).size * rank
                    rank = rank + 1
                End If
            End If
        Loop

        ' see if this gap is closer to the unit than the previous one
        xdiff = x - aDroid(psDroid).pos.x
        ydiff = y - aDroid(psDroid).pos.y
        dist = xdiff * xdiff + ydiff * ydiff
'        dist += psFormation->rankDist*psFormation->rankDist * rank*rank;
        If (dist < objDist And rank = chosenRank) Or rank < chosenRank Then
            ' this gap is nearer
            objDist = dist
'            objDist += psFormation->rankDist*psFormation->rankDist * rank*rank;

            chosenLine = line_
            chosenDist = currDist + objRadius
            chosenPrev = prev
            chosenRank = rank
            pX = x
            pY = y
        End If
    Next

    ' initialise the member
    unit = aFormation(psFormation).free
    aFormation(psFormation).free = aFormation(psFormation).asMembers(unit).next_
    aFormation(psFormation).asMembers(unit).line_ = ToInt8(chosenLine)
    aFormation(psFormation).asMembers(unit).dist = ToInt16(chosenDist)
    aFormation(psFormation).asMembers(unit).psDroid = psDroid

    ' insert the unit into the list
    If chosenPrev = -1 Then
        aFormation(psFormation).asMembers(unit).next_ = aFormation(psFormation).asLines(chosenLine).member
        aFormation(psFormation).asLines(chosenLine).member = ToInt8(unit)
    Else
        aFormation(psFormation).asMembers(unit).next_ = aFormation(psFormation).asMembers(chosenPrev).next_
        aFormation(psFormation).asMembers(chosenPrev).next_ = ToInt8(unit)
    End If
End Sub

' re-insert all the units in the formation
Private Sub formationReorder(ByVal psFormation As Long)
    Dim numObj As Long, i As Long, curr As Long, prev As Long
    Dim asDroids(0 To F_MAXMEMBERS - 1) As F_MEMBER
    Dim xdiff As Long, ydiff As Long, insert As Long
    Dim aDist(0 To F_MAXMEMBERS - 1) As Long
    Dim psDroid As Long

    ' first find all the units to insert
    numObj = 0
    For i = 0 To F_MAXMEMBERS - 1
        psDroid = aFormation(psFormation).asMembers(i).psDroid
        If psDroid <> 0 Then
            asDroids(numObj).psDroid = psDroid
            xdiff = aDroid(psDroid).pos.x - aFormation(psFormation).x
            ydiff = aDroid(psDroid).pos.y - aFormation(psFormation).y
            aDist(numObj) = xdiff * xdiff + ydiff * ydiff
            numObj = numObj + 1
        End If
    Next

    ' create a list in distance order
    insert = -1
    For i = 0 To numObj - 1
        If insert = -1 Then
            ' insert at the start of the list
            asDroids(i).next_ = -1
            insert = i
        Else
            prev = -1
            curr = insert
            Do While curr <> -1
                If aDist(i) < aDist(curr) Then
                    Exit Do
                End If
                prev = curr
                curr = asDroids(curr).next_
            Loop
            If prev = -1 Then
                ' insert at the start of the list
                asDroids(i).next_ = ToInt8(insert)
                insert = i
            Else
                asDroids(i).next_ = asDroids(prev).next_
                asDroids(prev).next_ = ToInt8(i)
            End If
        End If
    Next

    ' reset the free list
    aFormation(psFormation).free = 0
    For i = 0 To F_MAXMEMBERS - 1           ' memset(psFormation->asMembers, 0, ...)
        aFormation(psFormation).asMembers(i).line_ = 0
        aFormation(psFormation).asMembers(i).next_ = 0
        aFormation(psFormation).asMembers(i).dist = 0
        aFormation(psFormation).asMembers(i).psDroid = 0
    Next
    For i = 0 To F_MAXMEMBERS - 1
        aFormation(psFormation).asMembers(i).next_ = i + 1
    Next
    aFormation(psFormation).asMembers(F_MAXMEMBERS - 1).next_ = -1
    For i = 0 To aFormation(psFormation).numLines - 1
        aFormation(psFormation).asLines(i).member = -1
    Next

    ' insert each member again
    Do While insert <> -1
        formationFindFree psFormation, asDroids(insert).psDroid, xdiff, ydiff
        insert = asDroids(insert).next_
    Loop
End Sub

' get a target position to move into a formation
Public Function formationGetPos(ByVal psFormation As Long, ByVal psDroid As Long, pX As Long, pY As Long, ByVal bCheckLOS As Boolean) As Boolean
    Dim xdiff As Long, ydiff As Long, distSq As Long
    Dim member As Long, x As Long, y As Long

    If Not (psFormation <> 0) Then ASSERT_FAILURE "psFormation != nullptr", "src/formation.cpp:559", "formationGetPos", "Invalid formation pointer": formationGetPos = False: Exit Function

'    if (psFormation->refCount == 1)
'    {
'        // nothing else in the formation so don't do anything
'        return false;
'    }

    ' see if the unit is close enough to join the formation
    xdiff = aFormation(psFormation).x - aDroid(psDroid).pos.x
    ydiff = aFormation(psFormation).y - aDroid(psDroid).pos.y
    distSq = xdiff * xdiff + ydiff * ydiff
'    rangeSq = 3*psFormation->size/2;
'    rangeSq = rangeSq*rangeSq;
    If distSq > F_JOINRANGE * F_JOINRANGE Then
        formationGetPos = False
        Exit Function
    End If

    ' see if the unit is already in the formation
    For member = 0 To F_MAXMEMBERS - 1
        If aFormation(psFormation).asMembers(member).psDroid = psDroid Then
            Exit For
        End If
    Next

    If member < F_MAXMEMBERS Then
        ' the unit is already in the formation - return it's current position
        formationCalcPos psDroid, psFormation, aFormation(psFormation).asMembers(member).line_, aFormation(psFormation).asMembers(member).dist, x, y
    ElseIf aFormation(psFormation).free <> -1 Then
        ' add the new DROID to the members
        aFormation(psFormation).asMembers(aFormation(psFormation).free).psDroid = psDroid
        aFormation(psFormation).free = aFormation(psFormation).asMembers(aFormation(psFormation).free).next_
        formationReorder psFormation

        ' find the DROID
        For member = 0 To F_MAXMEMBERS - 1
            If aFormation(psFormation).asMembers(member).psDroid = psDroid Then
                Exit For
            End If
        Next

        ' calculate its position
        formationCalcPos psDroid, psFormation, aFormation(psFormation).asMembers(member).line_, aFormation(psFormation).asMembers(member).dist, x, y
    Else
        formationGetPos = False
        Exit Function
    End If

' NOTE: fpathTileLOS was removed in: https://github.com/Warzone2100/warzone2100/commit/1cc62d6d9d9d95354e10522c8618423565779b89
'    // check the unit can get to the formation position
'    if (bCheckLOS && !fpathTileLOS(psDroid, Vector3i_Init(x, y, 0)))
'    {
'        return false;
'    }

    pX = x
    pY = y

    formationGetPos = True
End Function

' See if a unit is a member of a formation (i.e. it has a position assigned)
Public Function formationMember(ByVal psFormation As Long, ByVal psDroid As Long) As Boolean
    Dim member As Long

    ' see if the unit is already in the formation
    For member = 0 To F_MAXMEMBERS - 1
        If aFormation(psFormation).asMembers(member).psDroid = psDroid Then
            formationMember = True
            Exit Function
        End If
    Next

    formationMember = False
End Function

Private Function formationObjRadius(ByVal psDroid As Long) As Long
    Dim psBdyStats As Long

    If Not (aDroid(psDroid).typ = OBJ_DROID) Then ASSERT_FAILURE "psDroid->type == OBJ_DROID", "src/formation.cpp:655", "formationObjRadius", "We got passed a DROID that isn't a DROID!"

    psBdyStats = aDroid(psDroid).asBits(COMP_BODY)

    Select Case asBodyStats(psBdyStats).size
    Case SIZE_LIGHT
        formationObjRadius = fmLtRad

    Case SIZE_MEDIUM
        formationObjRadius = fmMedRad

    Case SIZE_HEAVY
        formationObjRadius = fmHvyRad

    Case SIZE_SUPER_HEAVY
        formationObjRadius = 500

    Case Else
        formationObjRadius = 3 * iIMDBaseShape_radius(aDroid(psDroid).sDisplay.imd) \ 2
    End Select
End Function

' (VB) for the tests: the formations of a player, front first
Public Function formationListFirst(ByVal player As Long) As Long
    formationListFirst = fmHead(player)
End Function

Public Function formationListNext(ByVal psFormation As Long) As Long
    formationListNext = fmNext(psFormation)
End Function
