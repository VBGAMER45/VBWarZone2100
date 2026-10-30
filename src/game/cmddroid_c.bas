Attribute VB_Name = "cmddroid_c"
' src/cmddroid.cpp, cmddroid.h: Code for command droids.
Option Explicit

' This global instance is responsible for dealing with the each player's target designator.
Public apsCmdDesignator(0 To MAX_PLAYERS - 1) As Long    ' DROID *

' Last time the max commander limit message was displayed
Private lastMaxCmdLimitMsgTime As Long

Private Const MAX_COMMAND_LIMIT_MESSAGE_PAUSE As Long = 10000

' update the kills of a command droid if psKiller is in a command group
Public Sub cmdDroidUpdateExperience(ByVal psShooter As Long, ByVal experienceInc As Long)
    Dim psCommander As Long, room As Long

    If Not (psShooter <> 0) Then ASSERT_FAILURE "psShooter != nullptr", "src/cmddroid.cpp:200", "cmdDroidUpdateExperience", "invalid Unit pointer": Exit Sub

    If hasCommander(psShooter) Then
        If droidWithinCommanderRange(psShooter) Then
            psCommander = aGroup(aDroid(psShooter).psGroup).psCommander
            room = Not aDroid(psCommander).experience       ' UINT32_MAX - psCommander->experience
            If ULT(experienceInc, room) Then droidIncreaseExperience psCommander, experienceInc Else droidIncreaseExperience psCommander, room
        End If
    End If
End Sub

' This function returns true if the droid is assigned to a commander group and it is not the commander.
Public Function hasCommander(ByVal psDroid As Long) As Boolean
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/cmddroid.cpp:212", "hasCommander", "invalid droid pointer": hasCommander = False: Exit Function

    If aDroid(psDroid).droidType <> DROID_COMMAND And _
       aDroid(psDroid).psGroup <> 0 Then
        If aGroup(aDroid(psDroid).psGroup).typ = GT_COMMAND Then
            hasCommander = True
            Exit Function
        End If
    End If

    hasCommander = False
End Function

' This function returns the level of a droids commander. If the droid doesn't have commander, it returns 0.
Public Function cmdGetCommanderLevel(ByVal psDroid As Long) As Long
    Dim psCommander As Long

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/cmddroid.cpp:229", "cmdGetCommanderLevel", "invalid droid pointer"

    ' If this droid is not the member of a Commander's group
    ' Return an experience level of 0
    If Not hasCommander(psDroid) Then
        cmdGetCommanderLevel = 0
        Exit Function
    End If

    ' Retrieve this group's commander
    psCommander = aGroup(aDroid(psDroid).psGroup).psCommander

    ' Return the experience level of this commander
    cmdGetCommanderLevel = getDroidLevel(psCommander)
End Function
' This function allocs the global instance apsCmdDesignator.
Public Function cmdDroidInit() As Boolean
    Dim i As Long
    For i = 0 To MAX_PLAYERS - 1
        apsCmdDesignator(i) = 0
    Next
    cmdDroidInit = True
End Function

' ShutDown the command droids
Public Sub cmdDroidShutDown()
End Sub

' This function runs on all players to check if the player's current target designator as died.
' If it does, sets the target designator to NULL.
Public Sub cmdDroidUpdate()
    Dim i As Long
    For i = 0 To MAX_PLAYERS - 1
        If apsCmdDesignator(i) <> 0 Then
            If aDroid(apsCmdDesignator(i)).died <> 0 Then
                If Not (aDroid(apsCmdDesignator(i)).typ = OBJ_DROID) Then ASSERT_FAILURE "i->type == OBJ_DROID", "src/cmddroid.cpp:74", "cmdDroidUpdate", "Bad droid pointer! type=%u", aDroid(apsCmdDesignator(i)).typ
                apsCmdDesignator(i) = 0
            End If
        End If
    Next
End Sub

' This function adds the droid to the command group commanded by psCommander.
' It creates a group if it doesn't exist.
' If the group is not full, it adds the droid to it and sets all the droid's states and orders to the group's.
Public Function cmdDroidAddDroid(ByVal psCommander As Long, ByVal psDroid As Long) As Boolean
    Dim psGroup As Long, addedToGroup As Boolean, initialDroidGroup As Long, initialDroidRepairGroup As Long

    addedToGroup = False

    If Not (psCommander <> 0) Then ASSERT_FAILURE "psCommander != nullptr", "src/cmddroid.cpp:89", "cmdDroidAddDroid", "psCommander is null?": cmdDroidAddDroid = False: Exit Function
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/cmddroid.cpp:90", "cmdDroidAddDroid", "psDroid is null?": cmdDroidAddDroid = False: Exit Function

    initialDroidGroup = aDroid(psDroid).group
    initialDroidRepairGroup = aDroid(psDroid).repairGroup

    If aDroid(psCommander).psGroup = 0 Then
        psGroup = grpCreate()
        DROID_GROUP_add psGroup, psCommander
        aDroid(psDroid).repairGroup = UBYTE_MAX
        aDroid(psDroid).group = UBYTE_MAX
    End If

    If ULT(DROID_GROUP_getNumMembers(aDroid(psCommander).psGroup), cmdDroidMaxGroup(psCommander)) Then
        addedToGroup = True

        DROID_GROUP_add aDroid(psCommander).psGroup, psDroid
        aDroid(psDroid).repairGroup = UBYTE_MAX
        aDroid(psDroid).group = UBYTE_MAX

        ' set the secondary states for the unit
        ' dont reset DSO_ATTACK_RANGE, because there is no way to modify it under commander
        secondarySetState psDroid, DSO_REPAIR_LEVEL, aDroid(psCommander).secondaryOrder And DSS_REPLEV_MASK, ModeImmediate
        secondarySetState psDroid, DSO_ATTACK_LEVEL, aDroid(psCommander).secondaryOrder And DSS_ALEV_MASK, ModeImmediate
        secondarySetState psDroid, DSO_HALTTYPE, aDroid(psCommander).secondaryOrder And DSS_HALT_MASK, ModeImmediate

        orderDroidObj psDroid, DORDER_GUARD, psCommander, ModeImmediate
    ElseIf aDroid(psCommander).player = selectedPlayer Then
        'Do not potentially spam the console with this message
        If ULT(FromU32(CDbl(lastMaxCmdLimitMsgTime) + MAX_COMMAND_LIMIT_MESSAGE_PAUSE), gameTime) Then
            addConsoleMessage "Commander needs a higher level to command more units", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
            lastMaxCmdLimitMsgTime = gameTime
        End If

        If initialDroidGroup <> UBYTE_MAX Then
            aDroid(psDroid).group = initialDroidGroup
            SelectGroupDroid psDroid
        End If
        aDroid(psDroid).repairGroup = initialDroidRepairGroup
    End If

    If initialDroidGroup <> aDroid(psDroid).group Then
        intGroupsChanged
    End If

    cmdDroidAddDroid = addedToGroup
End Function

Public Function cmdDroidGetDesignator(ByVal player As Long) As Long
    cmdDroidGetDesignator = apsCmdDesignator(player)
End Function

Public Sub cmdDroidSetDesignator(ByVal psDroid As Long)
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/cmddroid.cpp:151", "cmdDroidSetDesignator", "Invalid droid!": Exit Sub
    If aDroid(psDroid).droidType <> DROID_COMMAND Then
        Exit Sub
    End If

    apsCmdDesignator(aDroid(psDroid).player) = psDroid
End Sub

Public Sub cmdDroidClearDesignator(ByVal player As Long)
    apsCmdDesignator(player) = 0
End Sub

' This function returns the index of the command droid.
' It does this by searching throughout all the player's droids.
' @todo try to find something more efficient, has this function is of O(TotalNumberOfDroidsOfPlayer).
Public Function cmdDroidGetIndex(ByVal psCommander As Long) As Long
    Dim index As Long, it As Long, psCurr As Long

    index = 1

    If aDroid(psCommander).droidType <> DROID_COMMAND Then
        cmdDroidGetIndex = 0
        Exit Function
    End If

    it = apsDroidLists(aDroid(psCommander).player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aDroid(psCurr).droidType = DROID_COMMAND And _
           ULT(aDroid(psCurr).id, aDroid(psCommander).id) Then
            index = index + 1
        End If
        it = lnNext(it)
    Loop

    cmdDroidGetIndex = index
End Function

' This function returns the maximum group size of the command droid.
Public Function cmdDroidMaxGroup(ByVal psCommander As Long) As Long
    Dim psStats As Long, player As Long
    psStats = DROID_getBrainStats(psCommander)
    player = aDroid(psCommander).player
    cmdDroidMaxGroup = FromU32(CDbl(getDroidLevel(psCommander)) * asBrainStats(psStats).upgrade(player).maxDroidsMult + asBrainStats(psStats).upgrade(player).maxDroids)
End Function
