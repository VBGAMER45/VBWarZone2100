Attribute VB_Name = "selection_c"
' src/selection.cpp, src/selection.h: selecting units by kind (the select-all keys), the next unit or
' building of a type, the n'th commander.
' (VB) selSelectUnitsIf's template over its condition is a condition number (SEL_*) and its value, run
' by selCondition. A std::vector<uint32_t> of components is its count and an array; the combinations'
' vectors are compared as their comma-joined text (equal vectors, equal text).
Option Explicit

Public Enum SELECTION_CLASS
    DS_ALL_UNITS
    DS_BY_TYPE
End Enum

Public Enum SELECTIONTYPE
    DST_UNUSED
    DST_VTOL
    DST_VTOL_ARMED
    DST_HOVER
    DST_WHEELED
    DST_TRACKED
    DST_HALF_TRACKED
    DST_CYBORG
    DST_ENGINEER
    DST_MECHANIC
    DST_TRANSPORTER
    DST_REPAIR_TANK
    DST_SENSOR
    DST_TRUCK
    DST_ALL_COMBAT
    DST_ALL_COMBAT_LAND
    DST_ALL_COMBAT_CYBORG
    DST_ALL_DAMAGED
    DST_ALL_SAME
    DST_ALL_LAND_MILDLY_OR_NOT_DAMAGED
End Enum

' (VB) selSelectUnitsIf's conditions
Private Const SEL_TRANSPORTER As Long = 0       ' selTransporter
Private Const SEL_TRUE As Long = 1              ' selTrue
Private Const SEL_PROP As Long = 2              ' selProp(droid, PROPULSION_TYPE)
Private Const SEL_PROP_ARMED As Long = 3        ' selPropArmed(droid, PROPULSION_TYPE)
Private Const SEL_TYPE As Long = 4              ' selType(droid, DROID_TYPE)
Private Const SEL_COMBAT As Long = 5            ' selCombat
Private Const SEL_COMBAT_LAND As Long = 6       ' selCombatLand
Private Const SEL_COMBAT_CYBORG As Long = 7     ' selCombatCyborg
Private Const SEL_DAMAGED As Long = 8           ' selDamaged
Private Const SEL_COMBAT_LAND_MILDLY_OR_NOT_DAMAGED As Long = 9 ' selCombatLandMildlyOrNotDamaged

' stores combinations of unit components
Private combinations() As String
Private combinations_size As Long

' selNextSpecifiedUnit's static DROID *psOldRD = nullptr; // pointer to last selected repair unit
Private psOldRD As Long
' selNextUnassignedUnit's static DROID *psOldNS = nullptr;
Private psOldNS As Long

' (VB) lib/framework/gettext.h's ngettext without translations: the singular for 1, else the plural
Private Function ngettext(ByVal msgid1 As String, ByVal msgid2 As String, ByVal n As Long) As String
    If n = 1 Then ngettext = msgid1 Else ngettext = msgid2
End Function

' (VB) selSelectUnitsIf's condition(psDroid) / condition(psDroid, value)
Private Function selCondition(ByVal condition As Long, ByVal value As Long, ByVal psDroid As Long) As Boolean
    Select Case condition
    Case SEL_TRANSPORTER: selCondition = selTransporter(psDroid)
    Case SEL_TRUE: selCondition = selTrue(psDroid)
    Case SEL_PROP: selCondition = selProp(psDroid, value)
    Case SEL_PROP_ARMED: selCondition = selPropArmed(psDroid, value)
    Case SEL_TYPE: selCondition = selType(psDroid, value)
    Case SEL_COMBAT: selCondition = selCombat(psDroid)
    Case SEL_COMBAT_LAND: selCondition = selCombatLand(psDroid)
    Case SEL_COMBAT_CYBORG: selCondition = selCombatCyborg(psDroid)
    Case SEL_DAMAGED: selCondition = selDamaged(psDroid)
    Case SEL_COMBAT_LAND_MILDLY_OR_NOT_DAMAGED: selCondition = selCombatLandMildlyOrNotDamaged(psDroid)
    End Select
End Function

' template <typename T> static unsigned selSelectUnitsIf(unsigned player, T condition, bool onlyOnScreen), and the
' (player, condition, value, onlyOnScreen) one (value unused by the one-argument conditions)
Private Function selSelectUnitsIf(ByVal player As Long, ByVal condition As Long, ByVal value As Long, ByVal onlyOnScreen As Boolean) As Long
    Dim count As Long, it As Long, psDroid As Long, shouldSelect As Boolean
    If ULT(MAX_PLAYERS - 1, player) Then selSelectUnitsIf = 0: Exit Function

    count = 0

    selDroidDeselect player

    ' Go through all.
    it = apsDroidLists(player).head
    Do While it <> 0
        psDroid = lnObj(it)
        shouldSelect = False
        If Not onlyOnScreen Then
            shouldSelect = selCondition(condition, value, psDroid)
        ElseIf objectOnScreen(psDroid, 0) Then
            shouldSelect = selCondition(condition, value, psDroid)
        End If
        If shouldSelect Then count = count + 1
        If shouldSelect And Not aDroid(psDroid).selected And (aDroid(psDroid).flags And BIT_(OBJECT_FLAG_UNSELECTABLE)) = 0 Then
            SelectDroid psDroid
        ElseIf Not shouldSelect And aDroid(psDroid).selected Then
            DeSelectDroid psDroid
        End If
        it = lnNext(it)
    Loop

    selSelectUnitsIf = count
End Function

Private Function selTransporter(ByVal droid_ As Long) As Boolean
    selTransporter = DROID_isTransporter(droid_)
End Function

Private Function selTrue(ByVal droid_ As Long) As Boolean
    selTrue = Not selTransporter(droid_)
End Function

Private Function selProp(ByVal droid_ As Long, ByVal prop As Long) As Boolean
    selProp = asPropulsionStats(DROID_getPropulsionStats(droid_)).propulsionType = prop And Not selTransporter(droid_)
End Function

Private Function selPropArmed(ByVal droid_ As Long, ByVal prop As Long) As Boolean
    selPropArmed = False
    If asPropulsionStats(DROID_getPropulsionStats(droid_)).propulsionType = prop Then
        If vtolFull(droid_) Then
            selPropArmed = Not selTransporter(droid_)
        End If
    End If
End Function

Private Function selType(ByVal droid_ As Long, ByVal typ As Long) As Boolean
    selType = aDroid(droid_).droidType = typ
End Function

Private Function selCombat(ByVal droid_ As Long) As Boolean
    selCombat = aDroid(droid_).asWeaps(0).nStat > 0 And Not selTransporter(droid_)
End Function

Private Function selCombatLand(ByVal droid_ As Long) As Boolean
    Dim typ As Long
    typ = asPropulsionStats(DROID_getPropulsionStats(droid_)).propulsionType
    selCombatLand = aDroid(droid_).asWeaps(0).nStat > 0 And (typ = PROPULSION_TYPE_WHEELED Or _
                                                            typ = PROPULSION_TYPE_HALF_TRACKED Or _
                                                            typ = PROPULSION_TYPE_TRACKED Or _
                                                            typ = PROPULSION_TYPE_HOVER Or _
                                                            typ = PROPULSION_TYPE_LEGGED)
End Function

Private Function selCombatCyborg(ByVal droid_ As Long) As Boolean
    Dim typ As Long
    typ = asPropulsionStats(DROID_getPropulsionStats(droid_)).propulsionType
    selCombatCyborg = aDroid(droid_).asWeaps(0).nStat > 0 And typ = PROPULSION_TYPE_LEGGED
End Function

' PERCENT(a, b) is math_ext.h's (((a)*100)/(b))
Private Function selDamaged(ByVal droid_ As Long) As Boolean
    selDamaged = (aDroid(droid_).body * 100) \ aDroid(droid_).originalBody < REPAIRLEV_LOW And Not selTransporter(droid_)
End Function

Private Function selNoGroup(ByVal psDroid As Long) As Boolean
    selNoGroup = aDroid(psDroid).group <> UBYTE_MAX
End Function

Private Function selCombatLandMildlyOrNotDamaged(ByVal psDroid As Long) As Boolean
    selCombatLandMildlyOrNotDamaged = False
    If (aDroid(psDroid).body * 100) \ aDroid(psDroid).originalBody > REPAIRLEV_LOW Then
        If selCombatLand(psDroid) Then
            selCombatLandMildlyOrNotDamaged = Not selNoGroup(psDroid)
        End If
    End If
End Function

' ---------------------------------------------------------------------
' Deselects all units for the player
Public Function selDroidDeselect(ByVal player As Long) As Long
    If vbDisplayOff Then selDroidDeselect = 0: Exit Function  ' (VB) the vbtest runs: statsref's stand-in
    Dim count As Long, it As Long, psDroid As Long
    count = 0
    If ULT(MAX_PLAYERS - 1, player) Then selDroidDeselect = 0: Exit Function

    it = apsDroidLists(player).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected Then
            count = count + 1
            DeSelectDroid psDroid
        End If
        it = lnNext(it)
    Loop

    selDroidDeselect = count
End Function

' ---------------------------------------------------------------------
' Lets you know how many are selected for a given player
Public Function selNumSelected(ByVal player As Long) As Long
    Dim count As Long, it As Long
    count = 0
    If ULT(MAX_PLAYERS - 1, player) Then selNumSelected = 0: Exit Function

    it = apsDroidLists(player).head
    Do While it <> 0
        If aDroid(lnObj(it)).selected Then
            count = count + 1
        End If
        it = lnNext(it)
    Loop

    selNumSelected = count
End Function

' std::vector<uint32_t> buildComponentsFromDroid(const DROID* psDroid): the count, the stats in components(0 .. count - 1)
Public Function buildComponentsFromDroid(ByVal psDroid As Long, components() As Long) As Long
    Dim stat As Long, c As Long, n As Long
    stat = 0
    n = 0
    ReDim components(0 To DROID_MAXCOMP + 1)

    ' stats are sorted by their estimated usefulness to distinguish units:
    ' * the first weapon turret is the most common difference between them
    '   followed by propulsions and bodies
    ' * the second weapon turret is least important because Hydras are rare
    For c = 0 To DROID_MAXCOMP + 1
        Select Case c
        Case 0: stat = aDroid(psDroid).asWeaps(1).nStat
        Case 1: stat = aDroid(psDroid).asBits(COMP_ECM)
        Case 2: stat = aDroid(psDroid).asBits(COMP_BRAIN)
        Case 3: stat = aDroid(psDroid).asBits(COMP_SENSOR)
        Case 4: stat = aDroid(psDroid).asBits(COMP_REPAIRUNIT)
        Case 5: stat = aDroid(psDroid).asBits(COMP_CONSTRUCT)
        Case 6: stat = aDroid(psDroid).asBits(COMP_BODY)
        Case 7: stat = aDroid(psDroid).asBits(COMP_PROPULSION)
        Case 8: stat = aDroid(psDroid).asWeaps(0).nStat
        End Select

        ' keep the list of components short by not adding stats with
        ' the value 0 to its end, since they are redundant
        If Not (stat = 0 And n = 0) Then
            components(n) = stat
            n = n + 1
        End If
    Next

    buildComponentsFromDroid = n
End Function

' (VB) a components vector as the text combinations holds
Private Function componentsKey(ByVal psDroid As Long) As String
    Dim components() As Long, n As Long, i As Long, s As String
    n = buildComponentsFromDroid(psDroid, components)
    s = ""
    For i = 0 To n - 1
        s = s & CStr(components(i)) & ","
    Next
    componentsKey = s
End Function

' Helper function to check whether the component stats of a unit can be found
' in the combinations vector and, optionally, to add them to it if not
Private Function componentsInCombinations(ByVal psDroid As Long, ByVal add As Boolean) As Boolean
    Dim components As String, it As Long
    components = componentsKey(psDroid)
    For it = 0 To combinations_size - 1
        If combinations(it) = components Then
            componentsInCombinations = True
            Exit Function
        End If
    Next
    ' add the list of components to the list of combinations unless
    ' this would result in a duplicate entry
    If add Then
        ReDim Preserve combinations(0 To combinations_size)
        combinations(combinations_size) = components
        combinations_size = combinations_size + 1
    End If
    componentsInCombinations = False
End Function

' Selects all units with the same propulsion, body and turret(s) as the one(s) selected
Private Function selSelectAllSame(ByVal player As Long, ByVal bOnScreen As Boolean) As Long
    Dim i As Long, selected As Long, excluded() As Long, excluded_size As Long, excluded_begin As Long
    Dim it As Long, psDroid As Long, onScreenExcluded As Boolean
    i = 0
    selected = 0
    excluded_size = 0
    excluded_begin = 0
    combinations_size = 0
    Erase combinations

    If ULT(MAX_PLAYERS - 1, player) Then selSelectAllSame = 0: Exit Function

    ' find out which units will need to be compared to which component combinations
    it = apsDroidLists(player).head
    Do While it <> 0
        psDroid = lnObj(it)
        onScreenExcluded = False
        If bOnScreen Then onScreenExcluded = Not objectOnScreen(psDroid, 0)
        If onScreenExcluded Then
            ReDim Preserve excluded(0 To excluded_size)
            excluded(excluded_size) = i
            excluded_size = excluded_size + 1
        ElseIf aDroid(psDroid).selected Then
            ReDim Preserve excluded(0 To excluded_size)
            excluded(excluded_size) = i
            excluded_size = excluded_size + 1
            selected = selected + 1
            componentsInCombinations psDroid, True
        End If
        i = i + 1
        it = lnNext(it)
    Loop

    ' if all or no units are selected, no more units can be chosen
    If excluded_size > 0 And i <> selected Then
        ' reset unit counter
        i = 0
        it = apsDroidLists(player).head
        Do While it <> 0
            psDroid = lnObj(it)
            ' excluded.empty() || *excluded.begin() != i (excluded.erase(excluded.begin()): excluded_begin moves on)
            If excluded_begin >= excluded_size Then
                If componentsInCombinations(psDroid, False) Then
                    SelectDroid psDroid
                    selected = selected + 1
                End If
            ElseIf excluded(excluded_begin) <> i Then
                If componentsInCombinations(psDroid, False) Then
                    SelectDroid psDroid
                    selected = selected + 1
                End If
            Else
                excluded_begin = excluded_begin + 1
            End If
            i = i + 1
            it = lnNext(it)
        Loop
    End If

    selSelectAllSame = selected
End Function

' ffs am
' ---------------------------------------------------------------------
Public Sub selNextSpecifiedUnit(ByVal unitType As Long)
    Dim psResult As Long, psFirst As Long, bLaterInList As Boolean, it As Long, psCurr As Long, bMatch As Boolean, resultAlive As Boolean
    psResult = 0
    psFirst = 0
    bLaterInList = False

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/selection.cpp:291", "selNextSpecifiedUnit", "invalid selectedPlayer: %u", selectedPlayer: Exit Sub

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psCurr = lnObj(it)
        'exceptions - as always...
        bMatch = False
        If unitType = DROID_CONSTRUCT Then
            If aDroid(psCurr).droidType = DROID_CONSTRUCT Or _
               aDroid(psCurr).droidType = DROID_CYBORG_CONSTRUCT Then
                bMatch = True
            End If
        ElseIf unitType = DROID_REPAIR Then
            If aDroid(psCurr).droidType = DROID_REPAIR Or _
               aDroid(psCurr).droidType = DROID_CYBORG_REPAIR Then
                bMatch = True
            End If
        ElseIf aDroid(psCurr).droidType = unitType Then
            bMatch = True
        End If
        If bMatch Then
            ' Always store away the first one we find
            If psFirst = 0 Then
                psFirst = psCurr
            End If

            If psCurr = psOldRD Then
                bLaterInList = True
            End If

            ' Nothing previously found...
            If psOldRD = 0 Then
                psResult = psCurr
                Exit Do
            ' Only select is this isn't the old one and it's further on in list
            ElseIf psCurr <> psOldRD And bLaterInList Then
                psResult = psCurr
                Exit Do
            End If
        End If
        it = lnNext(it)
    Loop

    ' Did we get one?
    If psResult = 0 And psFirst <> 0 Then
        psResult = psFirst
    End If

    resultAlive = False
    If psResult <> 0 Then resultAlive = aDroid(psResult).died = 0
    If resultAlive Then
        selDroidDeselect selectedPlayer
        SelectDroid psResult
        If getWarCamStatus() Then
            camToggleStatus ' messy - fix this
            ' setViewPos(map_coord(psCentreDroid->pos.x), map_coord(psCentreDroid->pos.y));
            processWarCam 'odd, but necessary
            camToggleStatus ' messy - FIXME
        Else
            ' camToggleStatus();
            ' Centre display on him if warcam isn't active
            setViewPos map_coord(aDroid(psResult).pos.x), map_coord(aDroid(psResult).pos.y), True
        End If
        psOldRD = psResult
    Else
        Select Case unitType
        Case DROID_REPAIR
            addConsoleMessage "Unable to locate any repair units!", LEFT_JUSTIFY, SYSTEM_MESSAGE
        Case DROID_CONSTRUCT
            addConsoleMessage "Unable to locate any Trucks!", LEFT_JUSTIFY, SYSTEM_MESSAGE
        Case DROID_SENSOR
            addConsoleMessage "Unable to locate any Sensor Units!", LEFT_JUSTIFY, SYSTEM_MESSAGE
        Case DROID_COMMAND
            addConsoleMessage "Unable to locate any Commanders!", LEFT_JUSTIFY, SYSTEM_MESSAGE
        Case Else
        End Select
    End If
End Sub

' ---------------------------------------------------------------------
Public Sub selNextUnassignedUnit()
    Dim psResult As Long, psFirst As Long, bLaterInList As Boolean, it As Long, psCurr As Long, resultAlive As Boolean
    psResult = 0
    psFirst = 0
    bLaterInList = False

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/selection.cpp:399", "selNextUnassignedUnit", "invalid selectedPlayer: %u", selectedPlayer: Exit Sub

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psCurr = lnObj(it)
        ' Only look at unselected ones
        If aDroid(psCurr).group = UBYTE_MAX Then
            ' Keep a record of first one
            If psFirst = 0 Then
                psFirst = psCurr
            End If

            If psCurr = psOldNS Then
                bLaterInList = True
            End If

            ' First one...?
            If psOldNS = 0 Then
                psResult = psCurr
                Exit Do
            ' Dont choose same one again
            ElseIf psCurr <> psOldNS And bLaterInList Then
                psResult = psCurr
                Exit Do
            End If
        End If
        it = lnNext(it)
    Loop

    ' If we didn't get one - then select first one
    If psResult = 0 And psFirst <> 0 Then
        psResult = psFirst
    End If

    resultAlive = False
    If psResult <> 0 Then resultAlive = aDroid(psResult).died = 0
    If resultAlive Then
        selDroidDeselect selectedPlayer
        SelectDroid psResult
        If getWarCamStatus() Then
            camToggleStatus ' messy - fix this
            ' setViewPos(map_coord(psCentreDroid->pos.x), map_coord(psCentreDroid->pos.y));
            processWarCam 'odd, but necessary
            camToggleStatus ' messy - FIXME
        Else
            ' camToggleStatus();
            ' Centre display on him if warcam isn't active
            setViewPos map_coord(aDroid(psResult).pos.x), map_coord(aDroid(psResult).pos.y), True
        End If
        psOldNS = psResult
    Else
        addConsoleMessage "Unable to locate any repair units!", LEFT_JUSTIFY, SYSTEM_MESSAGE
    End If
End Sub

' ---------------------------------------------------------------------
Public Sub selNextSpecifiedBuilding(ByVal structType As Long, ByVal jump As Boolean)
    Dim psResult As Long, psOldStruct As Long, psFirst As Long, bLaterInList As Boolean, it As Long, psCurr As Long
    Dim resultAlive As Boolean
    psResult = 0
    psOldStruct = 0
    psFirst = 0
    bLaterInList = False

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/selection.cpp:470", "selNextSpecifiedBuilding", "invalid selectedPlayer: %u", selectedPlayer: Exit Sub

    ' Firstly, start coughing if the type is invalid
    If ULT(NUM_DIFF_BUILDINGS, structType) Then ASSERT_FAILURE "structType <= NUM_DIFF_BUILDINGS", "src/selection.cpp:473", "selNextSpecifiedBuilding", "Invalid structure type %u", structType

    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        psCurr = lnObj(it)
        If psResult <> 0 Then
            Exit Do
        End If

        If asStructureStats(aStruct(psCurr).pStructureType).typ = structType And aStruct(psCurr).status = SS_BUILT Then
            If psFirst = 0 Then
                psFirst = psCurr
            End If

            If aStruct(psCurr).selected Then
                bLaterInList = True
                psOldStruct = psCurr
            ElseIf bLaterInList Then
                psResult = psCurr
            End If
        End If
        it = lnNext(it)
    Loop

    If psResult = 0 And psFirst <> 0 Then
        psResult = psFirst
    End If

    resultAlive = False
    If psResult <> 0 Then resultAlive = aStruct(psResult).died = 0
    If resultAlive Then
        If getWarCamStatus() Then
            camToggleStatus
        End If

        If jump Then
            setViewPos map_coord(aStruct(psResult).pos.x), map_coord(aStruct(psResult).pos.y), False
        End If

        If psOldStruct <> 0 Then
            aStruct(psOldStruct).selected = False
        End If
        aStruct(psResult).selected = True
        triggerEventSelected
        jsDebugSelected psResult
    Else
        ' Can't find required building
        addConsoleMessage "Cannot find required building!", LEFT_JUSTIFY, SYSTEM_MESSAGE
    End If
End Sub

' ---------------------------------------------------------------------
' see if a commander is the n'th command droid
Private Function droidIsCommanderNum(ByVal psDroid As Long, ByVal n As Long) As Boolean
    Dim numLess As Long, it As Long, psCurr As Long
    If aDroid(psDroid).droidType <> DROID_COMMAND Then
        droidIsCommanderNum = False
        Exit Function
    End If

    numLess = 0
    it = apsDroidLists(aDroid(psDroid).player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If (aDroid(psCurr).droidType = DROID_COMMAND) And ULT(aDroid(psCurr).id, aDroid(psDroid).id) Then
            numLess = numLess + 1
        End If
        it = lnNext(it)
    Loop

    droidIsCommanderNum = (numLess = (n - 1))
End Function

' select the n'th command droid
Public Sub selCommander(ByVal n As Long)
    Dim it As Long, psCurr As Long
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/selection.cpp:553", "selCommander", "invalid selectedPlayer: %u", selectedPlayer: Exit Sub

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psCurr = lnObj(it)
        If droidIsCommanderNum(psCurr, n) Then
            If Not aDroid(psCurr).selected And (aDroid(psCurr).flags And BIT_(OBJECT_FLAG_UNSELECTABLE)) = 0 Then
                clearSelection
                aDroid(psCurr).selected = True
            ElseIf (aDroid(psCurr).flags And BIT_(OBJECT_FLAG_UNSELECTABLE)) = 0 Then
                clearSelection

                aDroid(psCurr).selected = True

                ' this horrible bit of code is taken from activateGroupAndMove
                ' and sets the camera position to that of the commander

                If getWarCamStatus() Then
                    camToggleStatus ' messy - fix this
                    processWarCam ' odd, but necessary
                    camToggleStatus ' messy - FIXME
                Else
                    ' Centre display on him if warcam isn't active
                    setViewPos map_coord(aDroid(psCurr).pos.x), map_coord(aDroid(psCurr).pos.y), True
                End If
            End If
            Exit Sub
        End If
        it = lnNext(it)
    Loop
End Sub

' ---------------------------------------------------------------------
'
'   Selects the units of a given player according to given criteria.
'   It is also possible to request whether the units be onscreen or not.
'
Public Function selDroidSelection(ByVal player As Long, ByVal droidClass As SELECTION_CLASS, ByVal droidType As SELECTIONTYPE, ByVal bOnScreen As Boolean) As Long
    Dim retVal As Long
    If ULT(MAX_PLAYERS - 1, player) Then selDroidSelection = 0: Exit Function

    ' So far, we haven't selected any
    retVal = 0

    ' Establish the class of selection
    Select Case droidClass
    Case DS_ALL_UNITS
        retVal = selSelectUnitsIf(player, SEL_TRUE, 0, bOnScreen)
    Case DS_BY_TYPE
        Select Case droidType
        Case DST_VTOL
            retVal = selSelectUnitsIf(player, SEL_PROP, PROPULSION_TYPE_LIFT, bOnScreen)
        Case DST_VTOL_ARMED
            retVal = selSelectUnitsIf(player, SEL_PROP_ARMED, PROPULSION_TYPE_LIFT, bOnScreen)
        Case DST_HOVER
            retVal = selSelectUnitsIf(player, SEL_PROP, PROPULSION_TYPE_HOVER, bOnScreen)
        Case DST_WHEELED
            retVal = selSelectUnitsIf(player, SEL_PROP, PROPULSION_TYPE_WHEELED, bOnScreen)
        Case DST_TRACKED
            retVal = selSelectUnitsIf(player, SEL_PROP, PROPULSION_TYPE_TRACKED, bOnScreen)
        Case DST_HALF_TRACKED
            retVal = selSelectUnitsIf(player, SEL_PROP, PROPULSION_TYPE_HALF_TRACKED, bOnScreen)
        Case DST_CYBORG
            retVal = selSelectUnitsIf(player, SEL_PROP, PROPULSION_TYPE_LEGGED, bOnScreen)
        Case DST_ENGINEER
            retVal = selSelectUnitsIf(player, SEL_TYPE, DROID_CYBORG_CONSTRUCT, bOnScreen)
        Case DST_MECHANIC
            retVal = selSelectUnitsIf(player, SEL_TYPE, DROID_CYBORG_REPAIR, bOnScreen)
        Case DST_TRANSPORTER
            retVal = selSelectUnitsIf(player, SEL_TRANSPORTER, 0, bOnScreen)
        Case DST_REPAIR_TANK
            retVal = selSelectUnitsIf(player, SEL_TYPE, DROID_REPAIR, bOnScreen)
        Case DST_SENSOR
            retVal = selSelectUnitsIf(player, SEL_TYPE, DROID_SENSOR, bOnScreen)
        Case DST_TRUCK
            retVal = selSelectUnitsIf(player, SEL_TYPE, DROID_CONSTRUCT, bOnScreen)
        Case DST_ALL_COMBAT
            retVal = selSelectUnitsIf(player, SEL_COMBAT, 0, bOnScreen)
        Case DST_ALL_COMBAT_LAND
            retVal = selSelectUnitsIf(player, SEL_COMBAT_LAND, 0, bOnScreen)
        Case DST_ALL_COMBAT_CYBORG
            retVal = selSelectUnitsIf(player, SEL_COMBAT_CYBORG, 0, bOnScreen)
        Case DST_ALL_DAMAGED
            retVal = selSelectUnitsIf(player, SEL_DAMAGED, 0, bOnScreen)
        Case DST_ALL_SAME
            retVal = selSelectAllSame(player, bOnScreen)
        Case DST_ALL_LAND_MILDLY_OR_NOT_DAMAGED
            retVal = selSelectUnitsIf(player, SEL_COMBAT_LAND_MILDLY_OR_NOT_DAMAGED, 0, bOnScreen)
        Case Else
            ASSERT_FAILURE "false", "src/selection.cpp:668", "selDroidSelection", "Invalid selection type"
        End Select
    Case Else
        ASSERT_FAILURE "false", "src/selection.cpp:672", "selDroidSelection", "Invalid selection attempt"
    End Select

    CONPRINTF ngettext("%u unit selected", "%u units selected", retVal), retVal

    selDroidSelection = retVal
End Function
