Attribute VB_Name = "group_c"
' src/group.cpp, group.h: link droids together into a group for AI etc.
' A DROID_GROUP * is a 1-based handle into the pool aGroup (0 = nullptr);
' grpGlobalManager (std::map<int, unique_ptr<DROID_GROUP>>) is an array
' from the group id to its handle. Its methods are DROID_GROUP_<method>.
' The overloads of orderGroup: DROID_GROUP_orderGroup(this, order),
' DROID_GROUP_orderGroup_loc(.., x, y), DROID_GROUP_orderGroup_obj(.., psObj).
Option Explicit

Public Enum GROUP_TYPE
    GT_NORMAL       ' standard group
    GT_COMMAND      ' command droid group
    GT_TRANSPORTER  ' transporter group
End Enum

Public Type DROID_GROUP
    typ As GROUP_TYPE                   ' Type from the enum GROUP_TYPE above
    refCount As Long                    ' SWORD: Number of objects in the group. Group is deleted if refCount<=0. Count number of droids+NULL pointers.
    psList As OBJ_LIST                  ' DroidList: List of droids in the group
    psCommander As Long                 ' DROID *: The command droid of a command group
    id As Long                          ' unique group id
End Type

Public aGroup() As DROID_GROUP
Private aGroup_cap As Long
Private aGroup_free As Long             ' free handles, chained through id

' Group system variables: grpGlobalManager enables to remove all the groups to Shutdown the system
Private grpGlobalManager() As Long      ' group id -> handle, 0 = none
Private grpGlobalManager_cap As Long
Private grpInitialized As Boolean

' (VB) the unique_ptr's delete
Private Sub grpFree(ByVal psGroup As Long)
    objlist_clear aGroup(psGroup).psList
    aGroup(psGroup).id = aGroup_free
    aGroup_free = psGroup
End Sub

' (VB) grpGlobalManager.clear()
Private Sub grpGlobalManager_clear()
    Dim i As Long
    For i = 0 To grpGlobalManager_cap - 1
        If grpGlobalManager(i) <> 0 Then
            grpFree grpGlobalManager(i)
            grpGlobalManager(i) = 0
        End If
    Next
End Sub

' initialise the group system
Public Function grpInitialise() As Boolean
    grpGlobalManager_clear
    grpInitialized = True
    grpInitialise = True
End Function

' shutdown the group system
Public Sub grpShutDown()
    ' Since we are not very diligent removing groups after we have
    ' created them; we need this hack to remove them on level end.
    grpGlobalManager_clear
    grpInitialized = False
End Sub

' Constructor
Private Sub DROID_GROUP_construct(ByVal psGroup As Long)
    aGroup(psGroup).typ = GT_NORMAL
    aGroup(psGroup).refCount = 0
    aGroup(psGroup).psCommander = 0
End Sub

' create a new group
Public Function grpCreate() As Long
    Dim psGroup As Long, i As Long, n As Long
    If Not (grpInitialized) Then ASSERT_FAILURE "grpInitialized", "src/group.cpp:69", "grpCreate", "Group code not initialized yet"
    ' std::make_unique<DROID_GROUP>()
    If aGroup_free = 0 Then
        n = aGroup_cap
        If aGroup_cap = 0 Then aGroup_cap = 16 Else aGroup_cap = aGroup_cap * 2
        ReDim Preserve aGroup(1 To aGroup_cap)
        For i = aGroup_cap To n + 1 Step -1
            aGroup(i).id = aGroup_free
            aGroup_free = i
        Next
    End If
    psGroup = aGroup_free
    aGroup_free = aGroup(psGroup).id
    aGroup(psGroup).id = 0
    DROID_GROUP_construct psGroup

    i = 0
    Do While i < grpGlobalManager_cap   ' surly hack
        If grpGlobalManager(i) = 0 Then Exit Do
        i = i + 1
    Loop
    aGroup(psGroup).id = i
    If i = grpGlobalManager_cap Then
        If grpGlobalManager_cap = 0 Then grpGlobalManager_cap = 16 Else grpGlobalManager_cap = grpGlobalManager_cap * 2
        ReDim Preserve grpGlobalManager(0 To grpGlobalManager_cap - 1)
    End If
    grpGlobalManager(i) = psGroup
    grpCreate = psGroup
End Function

' (VB) psCommander != nullptr ? psCommander->id : 0
Private Function commanderId(ByVal psGroup As Long) As Long
    If aGroup(psGroup).psCommander <> 0 Then commanderId = aDroid(aGroup(psGroup).psCommander).id
End Function

' add a droid to a group
Public Sub DROID_GROUP_add(ByVal this As Long, ByVal psDroid As Long)
    If Not (grpInitialized) Then ASSERT_FAILURE "grpInitialized", "src/group.cpp:84", "add", "Group code not initialized yet"

    aGroup(this).refCount = ToInt16(aGroup(this).refCount + 1)
    ' if psDroid == NULL just increase the refcount don't add anything to the list
    If psDroid <> 0 Then
        If aGroup(this).psList.size <> 0 Then
            If aDroid(psDroid).player <> aDroid(objlist_front(aGroup(this).psList)).player Then
                ASSERT_FAILURE "false", "src/group.cpp:92", "add", "grpJoin: Cannot have more than one players droids in a group"
                Exit Sub
            End If
        End If

        If aDroid(psDroid).psGroup <> 0 Then
            DROID_GROUP_remove aDroid(psDroid).psGroup, psDroid
        End If
        aDroid(psDroid).psGroup = this

        If DROID_isTransporter(psDroid) Then
            If Not ((aGroup(this).typ = GT_NORMAL)) Then ASSERT_FAILURE "(type == GT_NORMAL)", "src/group.cpp:104", "add", "grpJoin: Cannot have two transporters in a group": Exit Sub
            If Not (aGroup(this).psList.size = 0) Then ASSERT_FAILURE "psList.empty()", "src/group.cpp:105", "add", "Adding transporter to non-empty list.": Exit Sub
            aGroup(this).typ = GT_TRANSPORTER
            objlist_push_front aGroup(this).psList, psDroid
        ElseIf (aDroid(psDroid).droidType = DROID_COMMAND) And (aGroup(this).typ <> GT_TRANSPORTER) Then
            If Not ((aGroup(this).typ = GT_NORMAL) And (aGroup(this).psCommander = 0)) Then ASSERT_FAILURE "(type == GT_NORMAL) && (psCommander == nullptr)", "src/group.cpp:111", "add", "grpJoin: Cannot have two command droids in a group": Exit Sub
            aGroup(this).typ = GT_COMMAND
            aGroup(this).psCommander = psDroid
        Else
            objlist_push_front aGroup(this).psList, psDroid
        End If

        If aGroup(this).typ = GT_COMMAND Then
            syncDebug "add", "Droid %d joining command group %d", aDroid(psDroid).id, commanderId(this)
        End If
    End If

    If aGroup(this).typ = GT_COMMAND And aGroup(this).psCommander <> 0 Then
        intCommanderGroupChanged aGroup(this).psCommander
    Else
        intCommanderGroupChanged 0
    End If
End Sub

' remove a droid from a group
Public Sub DROID_GROUP_remove(ByVal this As Long, ByVal psDroid As Long)
    Dim it As Long
    If Not (grpInitialized) Then ASSERT_FAILURE "grpInitialized", "src/group.cpp:132", "remove", "Group code not initialized yet": Exit Sub

    If psDroid <> 0 Then
        If aDroid(psDroid).psGroup <> this Then
            ASSERT_FAILURE "false", "src/group.cpp:136", "remove", "grpLeave: droid group does not match"
            Exit Sub
        End If
    End If

    ' SyncDebug
    If psDroid <> 0 And aGroup(this).typ = GT_COMMAND Then
        syncDebug "remove", "Droid %d leaving command group %d", aDroid(psDroid).id, commanderId(this)
    End If

    aGroup(this).refCount = ToInt16(aGroup(this).refCount - 1)
    ' if psDroid == NULL just decrease the refcount don't remove anything from the list
    If psDroid <> 0 Then
        ' update group list of droids
        If aDroid(psDroid).droidType <> DROID_COMMAND Or aGroup(this).typ <> GT_COMMAND Then
            it = objlist_find(aGroup(this).psList, psDroid)
            If it <> 0 Then
                objlist_erase aGroup(this).psList, it
            Else
                ASSERT_FAILURE "false", "src/group.cpp:160", "remove", "grpLeave: droid not found"
            End If
        End If

        aDroid(psDroid).psGroup = 0

        ' update group's type
        If (aDroid(psDroid).droidType = DROID_COMMAND) And (aGroup(this).typ = GT_COMMAND) Then
            aGroup(this).typ = GT_NORMAL
            aGroup(this).psCommander = 0
        ElseIf DROID_isTransporter(psDroid) And (aGroup(this).typ = GT_TRANSPORTER) Then
            aGroup(this).typ = GT_NORMAL
        End If
    End If

    If aGroup(this).typ = GT_COMMAND And aGroup(this).psCommander <> 0 Then
        intCommanderGroupChanged aGroup(this).psCommander
    Else
        intCommanderGroupChanged 0
    End If

    ' free the group if necessary
    If aGroup(this).refCount <= 0 Then
        grpGlobalManager(aGroup(this).id) = 0
        grpFree this
    End If
End Sub

' count the members of a group
Public Function DROID_GROUP_getNumMembers(ByVal this As Long) As Long
    If Not (grpInitialized) Then ASSERT_FAILURE "grpInitialized", "src/group.cpp:190", "getNumMembers", "Group code not initialized yet"
    DROID_GROUP_getNumMembers = aGroup(this).psList.size
End Function

' Give a group of droids an order
Public Sub DROID_GROUP_orderGroup(ByVal this As Long, ByVal order As DroidOrderType)
    Dim it As Long, psCurr As Long
    If Not (grpInitialized) Then ASSERT_FAILURE "grpInitialized", "src/group.cpp:197", "orderGroup", "Group code not initialized yet"

    it = aGroup(this).psList.head
    Do While it <> 0
        psCurr = lnObj(it)
        orderDroid psCurr, order, ModeQueue
        it = lnNext(it)
    Loop
End Sub

' Give a group of droids an order (using a Location)
Public Sub DROID_GROUP_orderGroup_loc(ByVal this As Long, ByVal order As DroidOrderType, ByVal x As Long, ByVal y As Long)
    Dim it As Long, psCurr As Long
    If Not (grpInitialized) Then ASSERT_FAILURE "grpInitialized", "src/group.cpp:208", "orderGroup", "Group code not initialized yet"
    If Not (validOrderForLoc(order)) Then ASSERT_FAILURE "validOrderForLoc(order)", "src/group.cpp:209", "orderGroup", "orderGroup: Bad order": Exit Sub

    it = aGroup(this).psList.head
    Do While it <> 0
        psCurr = lnObj(it)
        orderDroidLoc psCurr, order, x, y, IIf(bMultiMessages, ModeQueue, ModeImmediate)
        it = lnNext(it)
    Loop
End Sub

' Give a group of droids an order (using an Object)
Public Sub DROID_GROUP_orderGroup_obj(ByVal this As Long, ByVal order As DroidOrderType, ByVal psObj As Long)
    Dim it As Long, psCurr As Long
    If Not (validOrderForObj(order)) Then ASSERT_FAILURE "validOrderForObj(order)", "src/group.cpp:220", "orderGroup", "orderGroup: Bad order": Exit Sub

    it = aGroup(this).psList.head
    Do While it <> 0
        psCurr = lnObj(it)
        orderDroidObj psCurr, order, psObj, IIf(bMultiMessages, ModeQueue, ModeImmediate)
        it = lnNext(it)
    Loop
End Sub

' Set the secondary state for a group of droids
Public Sub DROID_GROUP_setSecondary(ByVal this As Long, ByVal sec As SECONDARY_ORDER, ByVal state As SECONDARY_STATE)
    Dim it As Long, psCurr As Long
    If Not (grpInitialized) Then ASSERT_FAILURE "grpInitialized", "src/group.cpp:231", "setSecondary", "Group code not initialized yet"

    it = aGroup(this).psList.head
    Do While it <> 0
        psCurr = lnObj(it)
        secondarySetState psCurr, sec, state
        it = lnNext(it)
    Loop
End Sub
