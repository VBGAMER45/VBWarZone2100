Attribute VB_Name = "objmem_c"
' src/objmem.cpp, objmem.h: the object lists and the object memory, the
' object ids, the flag positions, the lookups by id. The lists
' (PerPlayerObjectLists, std::list) are OBJ_LIST arrays (objlist_vb); a
' PerPlayer...Lists& parameter is the array, and `&pList[p] ==
' &apsDroidLists[p]` compares VarPtr. A FLAG_POSITION * is a 1-based
' handle into aFlagPos (calloc and free: a pool with a free list), 0 =
' nullptr. DEBUG-only code (objListIntegCheck, checkFactoryFlags) is left
' out, as in a release build.
Option Explicit

' the initial value for the object ID
Private Const OBJ_ID_INIT As Long = 20000

' The id number for the next object allocated
' Each object will have a unique id number irrespective of type
Public unsynchObjID As Long             ' uint32_t
Public synchObjID As Long               ' uint32_t

' The lists of objects allocated
Public apsDroidLists(0 To MAX_PLAYERS - 1) As OBJ_LIST
Public apsStructLists(0 To MAX_PLAYERS - 1) As OBJ_LIST
Public apsFeatureList(0 To 0) As OBJ_LIST
Public apsExtractorLists(0 To MAX_PLAYERS - 1) As OBJ_LIST
Public apsOilList(0 To 0) As OBJ_LIST
Public apsSensorList(0 To 0) As OBJ_LIST    ' List of sensors in the game.

'The list of Flag Positions allocated
Public apsFlagPosLists(0 To MAX_PLAYERS - 1) As OBJ_LIST

' The list of destroyed objects
Public psDestroyedObj As OBJ_LIST

' (VB) the FLAG_POSITION heap
Public aFlagPos() As FLAG_POSITION
Private aFlagPos_cap As Long
Private aFlagPos_free As Long           ' free handles, chained through factoryInc
Private flagPosEmpty As FLAG_POSITION

' Initialise the object heaps
Public Function objmemInitialise() As Boolean
    ' reset the object ID number
    unsynchObjID = OBJ_ID_INIT \ 2 ' /2 so that object IDs start around OBJ_ID_INIT*8, in case that's important when loading maps.
    synchObjID = OBJ_ID_INIT * 4   ' *4 so that object IDs start around OBJ_ID_INIT*8, in case that's important when loading maps.

    objmemInitialise = True
End Function

' Release the object heaps (objMemShutdownContainerImpl for each container)
Public Sub objmemShutdown()
    Dim h As Long
    h = DroidContainer_begin()
    Do While h <> 0
        ' Make sure to get rid of some final references in the sound code to this object first
        audio_RemoveObj h
        h = DroidContainer_next(h)
    Loop
    DroidContainer_clear
    h = StructContainer_begin()
    Do While h <> 0
        audio_RemoveObj h
        h = StructContainer_next(h)
    Loop
    StructContainer_clear
    h = FeatureContainer_begin()
    Do While h <> 0
        audio_RemoveObj h
        h = FeatureContainer_next(h)
    Loop
    FeatureContainer_clear
End Sub

Private Function objTypeToStr(ByVal typ As OBJECT_TYPE) As String
    Select Case typ
    Case OBJ_DROID: objTypeToStr = "droid"
    Case OBJ_STRUCTURE: objTypeToStr = "struct"
    Case OBJ_FEATURE: objTypeToStr = "feat"
    Case OBJ_PROJECTILE: objTypeToStr = "proj"
    Case Else: objTypeToStr = "object"
    End Select
End Function

Private Function getObjDebugDescriptiveName(ByVal psObj As Long) As String
    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        getObjDebugDescriptiveName = droidGetName(psObj)
    Case OBJ_STRUCTURE
        getObjDebugDescriptiveName = asStructureStats(aStruct(psObj).pStructureType).id
    Case OBJ_FEATURE
        getObjDebugDescriptiveName = asFeatureStats(aFeature(psObj).psStats).id
    Case Else
        getObjDebugDescriptiveName = "n/a"
    End Select
End Function

' (VB) the start of BADREF's message (release build)
Private Function badRefVictim(ByVal psVictim As Long) As String
    badRefVictim = va("Illegal reference to p%d:%s:%d (%s)", SIMPLE_OBJECT_player(psVictim), objTypeToStr(SIMPLE_OBJECT_type(psVictim)), SIMPLE_OBJECT_id(psVictim), getObjDebugDescriptiveName(psVictim))
End Function

' (VB) ASSERT_OR_RETURN_REPORT's failure: the assert, and the crash handler's report (not ported)
Private Sub assertReport(ByVal expr As String, ByVal line_ As Long, ByVal function_ As String, ByVal msg As String)
    ASSERT_FAILURE expr, "src/objmem.cpp:" & line_, function_, "%s", msg
End Sub

Private Function checkStructReferences_(ByVal psVictim As Long, psPlayerStructList As OBJ_LIST, ByVal player As Long, ByVal listName As String) As Boolean
    Dim it As Long, psStruct As Long, i As Long
    it = psPlayerStructList.head
    Do While it <> 0
        psStruct = lnObj(it)
        If psStruct = psVictim Then
            GoTo continue_  ' Don't worry about self references.
        End If

        For i = 0 To aStruct(psStruct).numWeaps - 1
            If Not (aStruct(psStruct).psTarget(i) <> psVictim) Then
                assertReport "psStruct->psTarget[i] != psVictim", 166, "_checkStructReferences", va("%s in %s[%u] (%s:%d - %s)", badRefVictim(psVictim), listName, player, objTypeToStr(aStruct(psStruct).typ), aStruct(psStruct).id, getObjDebugDescriptiveName(psStruct))
                Exit Function
            End If
        Next

        If aStruct(psStruct).pFunctionality <> 0 And aStruct(psStruct).pStructureType <> -1 Then
            Select Case asStructureStats(aStruct(psStruct).pStructureType).typ
            Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
                If Not (aStruct(psStruct).functionality.factory.psCommander <> psVictim) Then
                    assertReport "psFactory->psCommander != psVictim", 178, "_checkStructReferences", va("%s in FACTORY.psCommander in %s[%u]", badRefVictim(psVictim), listName, player)
                    Exit Function
                End If
            Case REF_POWER_GEN
                For i = 0 To NUM_POWER_MODULES - 1
                    If Not (aStruct(psStruct).functionality.powerGenerator.apResExtractors(i) <> psVictim) Then
                        assertReport "powerGen->apResExtractors[i] != psVictim", 186, "_checkStructReferences", va("%s in POWER_GEN.apResExtractors[%d] in %s[%u]", badRefVictim(psVictim), i, listName, player)
                        Exit Function
                    End If
                Next
            Case REF_RESOURCE_EXTRACTOR
                If Not (aStruct(psStruct).functionality.resourceExtractor.psPowerGen <> psVictim) Then
                    assertReport "psResExtracter->psPowerGen != psVictim", 193, "_checkStructReferences", va("%s in RES_EXTRACTOR.psPowerGen in %s[%u]", badRefVictim(psVictim), listName, player)
                    Exit Function
                End If
            Case REF_REPAIR_FACILITY
                If Not (aStruct(psStruct).functionality.repairFacility.psObj <> psVictim) Then
                    assertReport "psRepairFac->psObj != psVictim", 199, "_checkStructReferences", va("%s in REPAIR_FACILITY.psObj in %s[%u] (state=%d) (struct.status: %d)", badRefVictim(psVictim), listName, player, aStruct(psStruct).functionality.repairFacility.state, aStruct(psStruct).status)
                    Exit Function
                End If
            Case REF_REARM_PAD
                If Not (aStruct(psStruct).functionality.rearmPad.psObj <> psVictim) Then
                    assertReport "psReArmPad->psObj != psVictim", 205, "_checkStructReferences", va("%s in REARM_PAD.psObj in %s[%u] (struct.status: %d)", badRefVictim(psVictim), listName, player, aStruct(psStruct).status)
                    Exit Function
                End If
            End Select
        End If
continue_:
        it = lnNext(it)
    Loop
    checkStructReferences_ = True
End Function

Private Function checkDroidReferences_(ByVal psVictim As Long, psPlayerDroidList As OBJ_LIST, ByVal player As Long, ByVal listName As String) As Boolean
    Dim it As Long, psDroid As Long, i As Long
    it = psPlayerDroidList.head
    Do While it <> 0
        psDroid = lnObj(it)
        If psDroid = psVictim Then
            GoTo continue_  ' Don't worry about self references.
        End If

        If Not (aDroid(psDroid).order.psObj <> psVictim) Then
            assertReport "psDroid->order.psObj != psVictim", 225, "_checkDroidReferences", va("%s in order.psObj in %s[%u] (%s:%d - %s) (order.type: %s)", badRefVictim(psVictim), listName, player, objTypeToStr(aDroid(psDroid).typ), aDroid(psDroid).id, getObjDebugDescriptiveName(psDroid), getDroidOrderName(aDroid(psDroid).order.typ))
            Exit Function
        End If

        If Not (aDroid(psDroid).psBaseStruct <> psVictim) Then
            assertReport "psDroid->psBaseStruct != psVictim", 227, "_checkDroidReferences", va("%s in psBaseStruct in %s[%u] (%s:%d - %s)", badRefVictim(psVictim), listName, player, objTypeToStr(aDroid(psDroid).typ), aDroid(psDroid).id, getObjDebugDescriptiveName(psDroid))
            Exit Function
        End If

        For i = 0 To aDroid(psDroid).numWeaps - 1
            If aDroid(psDroid).psActionTarget(i) = psVictim Then
                assertReport "psDroid->psActionTarget[i] != psVictim", 233, "_checkDroidReferences", va("%s in %s[%u] (%s:%d - %s)", badRefVictim(psVictim), listName, player, objTypeToStr(aDroid(psDroid).typ), aDroid(psDroid).id, getObjDebugDescriptiveName(psDroid))
                Exit Function
            End If
        Next
continue_:
        it = lnNext(it)
    Loop
    checkDroidReferences_ = True
End Function

Private Function checkReferences(ByVal psVictim As Long) As Boolean
    Dim plr As Long
    For plr = 0 To MAX_PLAYERS - 1
        If Not checkStructReferences_(psVictim, apsStructLists(plr), plr, "apsStructLists") Then Exit Function
        If Not checkStructReferences_(psVictim, mission_apsStructLists(plr), plr, "mission.apsStructLists") Then Exit Function

        If Not checkDroidReferences_(psVictim, apsDroidLists(plr), plr, "apsDroidLists") Then Exit Function
        If Not checkDroidReferences_(psVictim, mission_apsDroidLists(plr), plr, "mission.apsDroidLists") Then Exit Function
        If Not checkDroidReferences_(psVictim, apsLimboDroids(plr), plr, "apsLimboDroids") Then Exit Function
    Next
    checkReferences = True
End Function

' Remove an object from the destroyed list, finally freeing its memory
' Hopefully by this time, no pointers still refer to it!
Public Function objmemDestroy(ByVal psObj As Long, ByVal checkRefs As Boolean) As Boolean
    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        debug_ 267, LOG_MEMORY, "objmemDestroy", "freeing droid at %p", psObj
    Case OBJ_STRUCTURE
        debug_ 271, LOG_MEMORY, "objmemDestroy", "freeing structure at %p", psObj
    Case OBJ_FEATURE
        debug_ 275, LOG_MEMORY, "objmemDestroy", "freeing feature at %p", psObj
    Case Else
        ASSERT_FAILURE "!""unknown object type""", "src/objmem.cpp:279", "objmemDestroy", "unknown object type in destroyed list at 0x%p", psObj
    End Select

    ' Make sure to get rid of some final references in the sound code to this object first
    audio_RemoveObj psObj

    If checkRefs Then
        If Not checkReferences(psObj) Then
            Exit Function
        End If
    End If
    If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        ' Droids are managed by a separate droid container.
        If Not (DroidContainer_find(psObj)) Then ASSERT_FAILURE "it != droidContainer.end()", "src/objmem.cpp:294", "objmemDestroy", "Droid not found in the global container!"
        DroidContainer_erase psObj
    ElseIf SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE Then
        ' Structs are managed by a separate struct container.
        If Not (StructContainer_find(psObj)) Then ASSERT_FAILURE "it != structContainer.end()", "src/objmem.cpp:302", "objmemDestroy", "Structure not found in the global container!"
        StructContainer_erase psObj
    ElseIf SIMPLE_OBJECT_type(psObj) = OBJ_FEATURE Then
        ' Features are managed by a separate feature container.
        If Not (FeatureContainer_find(psObj)) Then ASSERT_FAILURE "it != featureContainer.end()", "src/objmem.cpp:310", "objmemDestroy", "Feature not found in the global container!"
        FeatureContainer_erase psObj
    End If
    debug_ 317, LOG_MEMORY, "objmemDestroy", "BASE_OBJECT* is freed."
    objmemDestroy = True
End Function

' General housekeeping for the object system
Public Sub objmemUpdate()
    Dim it As Long
    ' Go through the destroyed objects list looking for objects that
    ' were destroyed before this turn

    ' First remove the objects from the start of the list
    it = psDestroyedObj.head
    Do While it <> 0
        If Not ULT(gameTime - deltaGameTime, SIMPLE_OBJECT_died(lnObj(it))) Then
            objmemDestroy lnObj(it), True
            it = objlist_erase(psDestroyedObj, it)
        Else
            ' do the object died callback
            triggerEventDestroyed lnObj(it)
            it = lnNext(it)
        End If
    Loop
End Sub

Public Function generateNewObjectId() As Long
    ' Generate even ID for unsynchronized objects. This is needed for debug objects, templates and other border lines cases that should preferably be removed one day.
    generateNewObjectId = unsynchObjID * MAX_PLAYERS * 2 + selectedPlayer * 2 ' Was taken from createObject, where 'player' was used instead of 'selectedPlayer'. Hope there are no stupid hacks that try to recover 'player' from the last 3 bits.
    unsynchObjID = unsynchObjID + 1
End Function

Public Function generateSynchronisedObjectId() As Long
    Dim ret As Long
    ' Generate odd ID for synchronized objects
    ret = synchObjID * 2 + 1
    synchObjID = synchObjID + 1
    syncDebug "generateSynchronisedObjectId", "New objectId = %u", ret
    generateSynchronisedObjectId = ret
End Function

' Add the object to its list
' \param list is a pointer to the object list
Private Sub addObjectToList(list() As OBJ_LIST, ByVal object As Long, ByVal player As Long)
    If Not (object <> 0) Then ASSERT_FAILURE "object != nullptr", "src/objmem.cpp:369", "addObjectToList", "Invalid pointer": Exit Sub
    If Not (player < UBound(list) + 1) Then ASSERT_FAILURE "player < PlayerCount", "src/objmem.cpp:370", "addObjectToList", "Invalid player index: %d", player: Exit Sub

    ' Prepend the object to the top of the list
    objlist_push_front list(player), object
End Sub

' Add the object to its list
' \param list is a pointer to the object list
Private Sub addObjectToFuncList(list() As OBJ_LIST, ByVal object As Long, ByVal player As Long)
    If Not (object <> 0) Then ASSERT_FAILURE "object != nullptr", "src/objmem.cpp:382", "addObjectToFuncList", "Invalid pointer": Exit Sub
    If Not (Not BASE_OBJECT_hasExtraFunction(object)) Then ASSERT_FAILURE "!object->hasExtraFunction", "src/objmem.cpp:383", "addObjectToFuncList", "%s(%p) is already in a function list!", objInfo(object), object: Exit Sub

    ' Prepend the object to the top of the list
    objlist_push_front list(player), object
    BASE_OBJECT_hasExtraFunction(object) = True
End Sub

' Move an object from the active list to the destroyed list.
' \param list is a pointer to the object list
' \param del is a pointer to the object to remove
Private Sub destroyObject(list() As OBJ_LIST, ByVal object As Long)
    Dim it As Long, player As Long
    If Not (object <> 0) Then ASSERT_FAILURE "object != nullptr", "src/objmem.cpp:397", "destroyObject", "Invalid pointer": Exit Sub
    player = SIMPLE_OBJECT_player(object)
    If Not (player < UBound(list) + 1) Then ASSERT_FAILURE "object->player < PlayerCount", "src/objmem.cpp:398", "destroyObject", "Invalid player index: %d", player: Exit Sub
    If Not (Not ULT(gameTime, gameTime - deltaGameTime) Or gameTime = 2) Then ASSERT_FAILURE "gameTime - deltaGameTime <= gameTime || gameTime == 2", "src/objmem.cpp:399", "destroyObject", "Expected %u <= %u, bad time", gameTime - deltaGameTime, gameTime

    it = objlist_find(list(player), object)
    If Not (it <> 0) Then ASSERT_FAILURE "it != list[object->player].end()", "src/objmem.cpp:402", "destroyObject", "Object %s(%d) not found in list", objInfo(object), SIMPLE_OBJECT_id(object)

    If it <> 0 Then
        objlist_erase list(player), it

        ' Prepend the object to the destruction list
        objlist_push_front psDestroyedObj, object

        ' Set destruction time
        SIMPLE_OBJECT_died(object) = gameTime
    End If
    scriptRemoveObject object
End Sub

' Remove an object from the active list
' \param list is a pointer to the object list
' \param remove is a pointer to the object to remove
' \param type is the type of the object
Private Sub removeObjectFromList(list() As OBJ_LIST, ByVal object As Long, ByVal player As Long)
    Dim it As Long
    If Not (object <> 0) Then ASSERT_FAILURE "object != nullptr", "src/objmem.cpp:425", "removeObjectFromList", "Invalid pointer": Exit Sub

    it = objlist_find(list(player), object)
    If Not (it <> 0) Then ASSERT_FAILURE "it != list[player].end()", "src/objmem.cpp:428", "removeObjectFromList", "Object %p not found in list", object: Exit Sub
    objlist_erase list(player), it
End Sub

' Remove an object from the relevant function list. An object can only be in one function list at a time!
' \param list is a pointer to the object list
' \param remove is a pointer to the object to remove
' \param type is the type of the object
Private Sub removeObjectFromFuncList(list() As OBJ_LIST, ByVal object As Long, ByVal player As Long)
    Dim it As Long
    If Not (object <> 0) Then ASSERT_FAILURE "object != nullptr", "src/objmem.cpp:440", "removeObjectFromFuncList", "Invalid pointer": Exit Sub

    ' Iterate through the list and find the item before the object to delete
    it = objlist_find(list(player), object)
    If Not (it <> 0) Then ASSERT_FAILURE "it != list[player].end()", "src/objmem.cpp:444", "removeObjectFromFuncList", "Object %p not found in list", object: Exit Sub
    objlist_erase list(player), it
    BASE_OBJECT_hasExtraFunction(object) = False
End Sub

'**************************************************************************************
'
' The actual object memory management functions for the different object types

'**************************  DROID  ********************************

' add the droid to the Droid Lists
Public Sub addDroid(ByVal psDroidToAdd As Long, pList() As OBJ_LIST)
    Dim psGroup As Long, player As Long

    addObjectToList pList, psDroidToAdd, aDroid(psDroidToAdd).player
    player = aDroid(psDroidToAdd).player

    ' Whenever a droid gets added to a list other than the current list
    ' its died flag is set to NOT_CURRENT_LIST so that anything targetting
    ' it will cancel itself - HACK?!
    If VarPtr(pList(player)) = VarPtr(apsDroidLists(player)) Then
        aDroid(psDroidToAdd).died = 0
        If aDroid(psDroidToAdd).droidType = DROID_SENSOR Then
            addObjectToFuncList apsSensorList, psDroidToAdd, 0
        End If

        ' commanders have to get their group back if not already loaded
        If aDroid(psDroidToAdd).droidType = DROID_COMMAND And aDroid(psDroidToAdd).psGroup = 0 Then
            psGroup = grpCreate()
            DROID_GROUP_add psGroup, psDroidToAdd
        End If
    ElseIf VarPtr(pList(player)) = VarPtr(mission_apsDroidLists(player)) Then
        If aDroid(psDroidToAdd).droidType = DROID_SENSOR Then
            addObjectToFuncList mission_apsSensorList, psDroidToAdd, 0
        End If
    End If
End Sub

' Destroy a droid
Public Sub killDroid(ByVal psDel As Long)
    Dim i As Long

    If Not (aDroid(psDel).typ = OBJ_DROID) Then ASSERT_FAILURE "psDel->type == OBJ_DROID", "src/objmem.cpp:514", "killDroid", "killUnit: pointer is not a unit"
    If Not (aDroid(psDel).player < MAX_PLAYERS) Then ASSERT_FAILURE "psDel->player < 11", "src/objmem.cpp:516", "killDroid", "killUnit: invalid player for unit"

    setDroidTarget psDel, 0
    For i = 0 To MAX_WEAPONS - 1
        setDroidActionTarget psDel, 0, i
    Next
    setDroidBase psDel, 0
    If aDroid(psDel).droidType = DROID_SENSOR Then
        removeObjectFromFuncList apsSensorList, psDel, 0
    End If

    destroyObject apsDroidLists, psDel
End Sub

' freeAllEntitiesImpl<DROID>
Private Sub freeAllDroidsImpl(entityLists() As OBJ_LIST)
    Dim l As Long, it As Long
    For l = 0 To UBound(entityLists)
        it = entityLists(l).head
        Do While it <> 0
            If Not DroidContainer_find(lnObj(it)) Then
                ASSERT_FAILURE "false", "src/objmem.cpp:595", "freeAllEntitiesImpl", "%s not found in the global container!", "Droid"
            Else
                DroidContainer_erase lnObj(it)
            End If
            it = lnNext(it)
        Loop
        objlist_clear entityLists(l)
    Next
End Sub

' Remove all droids
Public Sub freeAllDroids()
    freeAllDroidsImpl apsDroidLists
End Sub

'Remove a single Droid from a list
Public Sub removeDroid(ByVal psDroidToRemove As Long, pList() As OBJ_LIST)
    Dim player As Long
    If Not (aDroid(psDroidToRemove).typ = OBJ_DROID) Then ASSERT_FAILURE "psDroidToRemove->type == OBJ_DROID", "src/objmem.cpp:613", "removeDroid", "Pointer is not a unit": Exit Sub
    If Not (aDroid(psDroidToRemove).player < MAX_PLAYERS) Then ASSERT_FAILURE "psDroidToRemove->player < MAX_PLAYERS", "src/objmem.cpp:614", "removeDroid", "Invalid player for unit": Exit Sub
    player = aDroid(psDroidToRemove).player
    removeObjectFromList pList, psDroidToRemove, player

    ' Whenever a droid is removed from the current list its died
    ' flag is set to NOT_CURRENT_LIST so that anything targetting
    ' it will cancel itself, and we know it is not really on the map.
    If VarPtr(pList(player)) = VarPtr(apsDroidLists(player)) Then
        If aDroid(psDroidToRemove).droidType = DROID_SENSOR Then
            removeObjectFromFuncList apsSensorList, psDroidToRemove, 0
        End If
        aDroid(psDroidToRemove).died = NOT_CURRENT_LIST
    ElseIf VarPtr(pList(player)) = VarPtr(mission_apsDroidLists(player)) Then
        If aDroid(psDroidToRemove).droidType = DROID_SENSOR Then
            removeObjectFromFuncList mission_apsSensorList, psDroidToRemove, 0
        End If
    End If
End Sub

'Removes all droids that may be stored in the mission lists
Public Sub freeAllMissionDroids()
    freeAllDroidsImpl mission_apsDroidLists
End Sub

'Removes all droids that may be stored in the limbo lists
Public Sub freeAllLimboDroids()
    freeAllDroidsImpl apsLimboDroids
End Sub

'*************************  STRUCTURE  ******************************

' (VB) psStructure->pStructureType->pSensor && psStructure->pStructureType->pSensor->location == LOC_TURRET
Private Function hasTurretSensor(ByVal psStructure As Long) As Boolean
    Dim pSensor As Long
    pSensor = asStructureStats(aStruct(psStructure).pStructureType).pSensor
    If pSensor <> -1 Then hasTurretSensor = asSensorStats(pSensor).location = LOC_TURRET
End Function

' add the structure to the Structure Lists
Public Sub addStructure(ByVal psStructToAdd As Long)
    addObjectToList apsStructLists, psStructToAdd, aStruct(psStructToAdd).player
    If hasTurretSensor(psStructToAdd) Then
        addObjectToFuncList apsSensorList, psStructToAdd, 0
    ElseIf asStructureStats(aStruct(psStructToAdd).pStructureType).typ = REF_RESOURCE_EXTRACTOR Then
        addObjectToFuncList apsExtractorLists, psStructToAdd, aStruct(psStructToAdd).player
    End If
End Sub

' Destroy a structure
Public Sub killStruct(ByVal psBuilding As Long)
    Dim i As Long

    If Not (aStruct(psBuilding).typ = OBJ_STRUCTURE) Then ASSERT_FAILURE "psBuilding->type == OBJ_STRUCTURE", "src/objmem.cpp:671", "killStruct", "killStruct: pointer is not a droid"
    If Not (aStruct(psBuilding).player < MAX_PLAYERS) Then ASSERT_FAILURE "psBuilding->player < 11", "src/objmem.cpp:673", "killStruct", "killStruct: invalid player for structure"

    If hasTurretSensor(psBuilding) Then
        removeObjectFromFuncList apsSensorList, psBuilding, 0
    ElseIf asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_RESOURCE_EXTRACTOR Then
        removeObjectFromFuncList apsExtractorLists, psBuilding, aStruct(psBuilding).player
    End If

    For i = 0 To MAX_WEAPONS - 1
        setStructureTarget psBuilding, 0, i, ORIGIN_UNKNOWN
    Next

    If aStruct(psBuilding).pFunctionality <> 0 Then
        If STRUCTURE_isFactory(psBuilding) Then
            ' remove any commander from the factory
            If aStruct(psBuilding).functionality.factory.psCommander <> 0 Then
                assignFactoryCommandDroid psBuilding, 0
            End If

            ' remove any assembly points
            If aStruct(psBuilding).functionality.factory.psAssemblyPoint <> 0 Then
                removeFlagPosition aStruct(psBuilding).functionality.factory.psAssemblyPoint
                aStruct(psBuilding).functionality.factory.psAssemblyPoint = 0
            End If
        ElseIf asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_REPAIR_FACILITY Then
            If aStruct(psBuilding).functionality.repairFacility.psDeliveryPoint <> 0 Then
                ' free up repair fac stuff
                removeFlagPosition aStruct(psBuilding).functionality.repairFacility.psDeliveryPoint
                aStruct(psBuilding).functionality.repairFacility.psDeliveryPoint = 0
            End If
        End If
    End If

    destroyObject apsStructLists, psBuilding
End Sub

' Remove heapall structures (freeAllEntitiesImpl<STRUCTURE>)
Public Sub freeAllStructs()
    Dim l As Long, it As Long
    For l = 0 To MAX_PLAYERS - 1
        it = apsStructLists(l).head
        Do While it <> 0
            If Not StructContainer_find(lnObj(it)) Then
                ASSERT_FAILURE "false", "src/objmem.cpp:595", "freeAllEntitiesImpl", "%s not found in the global container!", "Structure"
            Else
                StructContainer_erase lnObj(it)
            End If
            it = lnNext(it)
        Loop
        objlist_clear apsStructLists(l)
    Next
End Sub

'Remove a single Structure from a list
Public Sub removeStructureFromList(ByVal psStructToRemove As Long, pList() As OBJ_LIST)
    If Not (aStruct(psStructToRemove).typ = OBJ_STRUCTURE) Then ASSERT_FAILURE "psStructToRemove->type == OBJ_STRUCTURE", "src/objmem.cpp:735", "removeStructureFromList", "removeStructureFromList: pointer is not a structure"
    If Not (aStruct(psStructToRemove).player < MAX_PLAYERS) Then ASSERT_FAILURE "psStructToRemove->player < 11", "src/objmem.cpp:737", "removeStructureFromList", "removeStructureFromList: invalid player for structure"
    removeObjectFromList pList, psStructToRemove, aStruct(psStructToRemove).player
    If hasTurretSensor(psStructToRemove) Then
        removeObjectFromFuncList apsSensorList, psStructToRemove, 0
    ElseIf asStructureStats(aStruct(psStructToRemove).pStructureType).typ = REF_RESOURCE_EXTRACTOR Then
        removeObjectFromFuncList apsExtractorLists, psStructToRemove, aStruct(psStructToRemove).player
    End If
End Sub

'*************************  FEATURE  ********************************

' add the feature to the Feature Lists
Public Sub addFeature(ByVal psFeatureToAdd As Long)
    addObjectToList apsFeatureList, psFeatureToAdd, 0
    If asFeatureStats(aFeature(psFeatureToAdd).psStats).subType = FEAT_OIL_RESOURCE Then
        addObjectToFuncList apsOilList, psFeatureToAdd, 0
    End If
End Sub

' Destroy a feature
' set the player to 0 since features have player = maxplayers+1. This screws up destroyObject
' it's a bit of a hack, but hey, it works
Public Sub killFeature(ByVal psDel As Long)
    If Not (aFeature(psDel).typ = OBJ_FEATURE) Then ASSERT_FAILURE "psDel->type == OBJ_FEATURE", "src/objmem.cpp:768", "killFeature", "killFeature: pointer is not a feature"
    aFeature(psDel).player = 0
    destroyObject apsFeatureList, psDel

    If asFeatureStats(aFeature(psDel).psStats).subType = FEAT_OIL_RESOURCE Then
        removeObjectFromFuncList apsOilList, psDel, 0
    End If
End Sub

' Remove all features (freeAllEntitiesImpl<FEATURE, 1>)
Public Sub freeAllFeatures()
    Dim it As Long
    it = apsFeatureList(0).head
    Do While it <> 0
        If Not FeatureContainer_find(lnObj(it)) Then
            ASSERT_FAILURE "false", "src/objmem.cpp:595", "freeAllEntitiesImpl", "%s not found in the global container!", "Feature"
        Else
            FeatureContainer_erase lnObj(it)
        End If
        it = lnNext(it)
    Loop
    objlist_clear apsFeatureList(0)
End Sub

'*************************  FLAG_POSITION *******************************

' (VB) calloc(1, sizeof(FLAG_POSITION))
Private Function flagPosCalloc() As Long
    Dim i As Long, n As Long
    If aFlagPos_free = 0 Then
        n = aFlagPos_cap
        If aFlagPos_cap = 0 Then aFlagPos_cap = 64 Else aFlagPos_cap = aFlagPos_cap * 2
        ReDim Preserve aFlagPos(1 To aFlagPos_cap)
        For i = aFlagPos_cap To n + 1 Step -1
            aFlagPos(i).factoryInc = aFlagPos_free
            aFlagPos_free = i
        Next
    End If
    flagPosCalloc = aFlagPos_free
    aFlagPos_free = aFlagPos(aFlagPos_free).factoryInc
    aFlagPos(flagPosCalloc) = flagPosEmpty
End Function

' (VB) free(psFlag)
Private Sub flagPosFree(ByVal psFlag As Long)
    aFlagPos(psFlag).factoryInc = aFlagPos_free
    aFlagPos_free = psFlag
End Sub

' Create a new Flag Position
Public Function createFlagPosition(ppsNew As Long, ByVal player As Long) As Boolean
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < 11", "src/objmem.cpp:790", "createFlagPosition", "createFlagPosition: invalid player number"

    ppsNew = flagPosCalloc()
    aFlagPos(ppsNew).typ = POS_DELIVERY
    aFlagPos(ppsNew).player = player
    aFlagPos(ppsNew).frameNumber = 0
    aFlagPos(ppsNew).selected = False
    aFlagPos(ppsNew).coords.x = Not 0
    aFlagPos(ppsNew).coords.y = Not 0
    aFlagPos(ppsNew).coords.z = Not 0
    createFlagPosition = True
End Function

Private Function isFlagPositionInList(ByVal psFlagPosToAdd As Long, list() As OBJ_LIST) As Boolean
    If Not (psFlagPosToAdd <> 0) Then ASSERT_FAILURE "psFlagPosToAdd != nullptr", "src/objmem.cpp:810", "isFlagPositionInList", "Invalid FlagPosition pointer": Exit Function
    If Not (ULT(aFlagPos(psFlagPosToAdd).player, MAX_PLAYERS)) Then ASSERT_FAILURE "psFlagPosToAdd->player < MAX_PLAYERS", "src/objmem.cpp:811", "isFlagPositionInList", "Invalid FlagPosition player: %u", aFlagPos(psFlagPosToAdd).player: Exit Function
    isFlagPositionInList = objlist_find(list(aFlagPos(psFlagPosToAdd).player), psFlagPosToAdd) <> 0
End Function

' add the Flag Position to the Flag Position Lists
Public Sub addFlagPositionToList(ByVal psFlagPosToAdd As Long, list() As OBJ_LIST)
    If Not (psFlagPosToAdd <> 0) Then ASSERT_FAILURE "psFlagPosToAdd != nullptr", "src/objmem.cpp:819", "addFlagPositionToList", "Invalid FlagPosition pointer": Exit Sub
    If Not (aFlagPos(psFlagPosToAdd).coords.x <> Not 0) Then ASSERT_FAILURE "psFlagPosToAdd->coords.x != ~0", "src/objmem.cpp:820", "addFlagPositionToList", "flag has invalid position": Exit Sub
    If Not (ULT(aFlagPos(psFlagPosToAdd).player, MAX_PLAYERS)) Then ASSERT_FAILURE "psFlagPosToAdd->player < MAX_PLAYERS", "src/objmem.cpp:821", "addFlagPositionToList", "Invalid FlagPosition player: %u", aFlagPos(psFlagPosToAdd).player: Exit Sub
    If isFlagPositionInList(psFlagPosToAdd, list) Then
        debug_ 824, LOG_INFO, "addFlagPositionToList", "FlagPosition is already in the list - ignoring"
        Exit Sub
    End If
    objlist_push_front list(aFlagPos(psFlagPosToAdd).player), psFlagPosToAdd
End Sub

' add the Flag Position to the Flag Position Lists
Public Sub addFlagPosition(ByVal psFlagPosToAdd As Long)
    addFlagPositionToList psFlagPosToAdd, apsFlagPosLists
End Sub

' Remove it from the list, but don't delete it!
Private Function removeFlagPositionFromList(ByVal psRemove As Long) As Boolean
    Dim it As Long
    If Not (psRemove <> 0) Then ASSERT_FAILURE "psRemove != nullptr", "src/objmem.cpp:839", "removeFlagPositionFromList", "Invalid Flag Position pointer": Exit Function
    If Not (ULT(aFlagPos(psRemove).player, MAX_PLAYERS)) Then ASSERT_FAILURE "psRemove->player < MAX_PLAYERS", "src/objmem.cpp:840", "removeFlagPositionFromList", "Invalid Flag Position player: %u", aFlagPos(psRemove).player: Exit Function

    it = objlist_find(apsFlagPosLists(aFlagPos(psRemove).player), psRemove)
    If it <> 0 Then
        objlist_erase apsFlagPosLists(aFlagPos(psRemove).player), it
        removeFlagPositionFromList = True
    End If
End Function

' Remove a Flag Position from the Lists
Public Sub removeFlagPosition(ByVal psDel As Long)
    If Not (psDel <> 0) Then ASSERT_FAILURE "psDel != nullptr", "src/objmem.cpp:856", "removeFlagPosition", "Invalid Flag Position pointer": Exit Sub

    If removeFlagPositionFromList(psDel) Then
        flagPosFree psDel
    Else
        ASSERT_FAILURE "false", "src/objmem.cpp:864", "removeFlagPosition", "Did not find flag position in expected list?"
    End If
End Sub

' Transfer a Flag Position to a new player
Public Sub transferFlagPositionToPlayer(ByVal psFlagPos As Long, ByVal originalPlayer As Long, ByVal newPlayer As Long)
    If Not (psFlagPos <> 0) Then ASSERT_FAILURE "psFlagPos != nullptr", "src/objmem.cpp:871", "transferFlagPositionToPlayer", "Invalid Flag Position pointer": Exit Sub
    If Not (originalPlayer = aFlagPos(psFlagPos).player) Then ASSERT_FAILURE "originalPlayer == psFlagPos->player", "src/objmem.cpp:872", "transferFlagPositionToPlayer", "Unexpected originalPlayer (%u) does not match current flagPos->player (%u)", originalPlayer, aFlagPos(psFlagPos).player
    If Not (removeFlagPositionFromList(psFlagPos)) Then ASSERT_FAILURE "removeFlagPositionFromList(psFlagPos)", "src/objmem.cpp:873", "transferFlagPositionToPlayer", "Did not find flag position in expected list?"
    aFlagPos(psFlagPos).player = newPlayer
    addFlagPosition psFlagPos
End Sub

' free all flag positions
Public Sub freeAllFlagPositions()
    Dim player As Long, it As Long
    For player = 0 To MAX_PLAYERS - 1
        it = apsFlagPosLists(player).head
        Do While it <> 0
            If Not (player = aFlagPos(lnObj(it)).player) Then ASSERT_FAILURE "player == flagPos->player", "src/objmem.cpp:885", "freeAllFlagPositions", "Player mismatch? (flagPos->player == %u, expecting: %d", aFlagPos(lnObj(it)).player, player
            flagPosFree lnObj(it)
            it = lnNext(it)
        Loop
        objlist_clear apsFlagPosLists(player)
    Next
End Sub

'*************************  OBJECT ACCESS FUNCTIONALITY *******************************

Private Function getBaseObjFromDroidId(list As OBJ_LIST, ByVal id As Long) As Long
    Dim it As Long, psObj As Long, it2 As Long
    it = list.head
    Do While it <> 0
        psObj = lnObj(it)
        If aDroid(psObj).id = id Then
            getBaseObjFromDroidId = psObj
            Exit Function
        End If
        ' if transporter check any droids in the grp
        If (aDroid(psObj).typ = OBJ_DROID) And DROID_isTransporter(psObj) Then
            If Not (aDroid(psObj).psGroup <> 0) Then ASSERT_FAILURE "psObj->psGroup != nullptr", "src/objmem.cpp:941", "getBaseObjFromDroidId", "Transporter has null group?": Exit Function
            it2 = aGroup(aDroid(psObj).psGroup).psList.head
            Do While it2 <> 0
                If aDroid(lnObj(it2)).id = id Then
                    getBaseObjFromDroidId = lnObj(it2)
                    Exit Function
                End If
                it2 = lnNext(it2)
            Loop
        End If
        it = lnNext(it)
    Loop
End Function

' getBaseObjFromId(const std::list<ObjectType*>& list, unsigned id)
Public Function getBaseObjFromId_list(list As OBJ_LIST, ByVal id As Long) As Long
    Dim it As Long
    it = list.head
    Do While it <> 0
        If SIMPLE_OBJECT_id(lnObj(it)) = id Then
            getBaseObjFromId_list = lnObj(it)
            Exit Function
        End If
        it = lnNext(it)
    Loop
End Function

' Find a base object from its id
Public Function getBaseObjFromData(ByVal id As Long, ByVal player As Long, ByVal typ As OBJECT_TYPE) As Long
    Dim p As Long
    If Not (ULT(player, MAX_PLAYERS) Or typ = OBJ_FEATURE) Then ASSERT_FAILURE "player < MAX_PLAYERS || type == OBJ_FEATURE", "src/objmem.cpp:957", "getBaseObjFromData", "Invalid player: %u", player: Exit Function

    Select Case typ
    Case OBJ_DROID
        p = getBaseObjFromDroidId(apsDroidLists(player), id)
        If p <> 0 Then getBaseObjFromData = p: Exit Function
        p = getBaseObjFromDroidId(mission_apsDroidLists(player), id)
        If p <> 0 Then getBaseObjFromData = p: Exit Function
        If player = 0 Then
            p = getBaseObjFromDroidId(apsLimboDroids(0), id)
            If p <> 0 Then getBaseObjFromData = p: Exit Function
        End If
    Case OBJ_STRUCTURE
        p = getBaseObjFromId_list(apsStructLists(player), id)
        If p <> 0 Then getBaseObjFromData = p: Exit Function
        p = getBaseObjFromId_list(mission_apsStructLists(player), id)
        If p <> 0 Then getBaseObjFromData = p: Exit Function
    Case OBJ_FEATURE
        p = getBaseObjFromId_list(apsFeatureList(0), id)
        If p <> 0 Then getBaseObjFromData = p: Exit Function
        p = getBaseObjFromId_list(mission_apsFeatureList(0), id)
        If p <> 0 Then getBaseObjFromData = p: Exit Function
    End Select
End Function

' Find a base object from it's id
Public Function getBaseObjFromId(ByVal id As Long) As Long
    Dim typ As Long, player As Long, psObj As Long
    ' Only cover OBJ_DROID, OBJ_STRUCTURE and OBJ_FEATURE types
    For typ = OBJ_DROID To OBJ_PROJECTILE - 1
        For player = 0 To MAX_PLAYERS - 1
            psObj = getBaseObjFromData(id, player, typ)
            If psObj <> 0 Then
                getBaseObjFromId = psObj
                Exit Function
            End If
        Next
    Next
    ASSERT_FAILURE "!""couldn't find a BASE_OBJ with ID""", "src/objmem.cpp:1032", "getBaseObjFromId", "getBaseObjFromId() failed for id %d", id
End Function

Private Function getRepairIdFromFlagSingleList(ByVal psFlag As Long, ByVal player As Long, list As OBJ_LIST) As Long
    Dim it As Long, psObj As Long
    it = list.head
    Do While it <> 0
        psObj = lnObj(it)
        If aStruct(psObj).pFunctionality <> 0 Then
            If asStructureStats(aStruct(psObj).pStructureType).typ = REF_REPAIR_FACILITY Then
                'check for matching delivery point
                ' ((REPAIR_FACILITY*)psObj->pFunctionality): the union's first member is at the same address
                If aStruct(psObj).functionality.repairFacility.psDeliveryPoint = psFlag Then
                    getRepairIdFromFlagSingleList = aStruct(psObj).id
                    Exit Function
                End If
            End If
        End If
        it = lnNext(it)
    Loop
    getRepairIdFromFlagSingleList = UDWORD_MAX
End Function

Public Function getRepairIdFromFlag(ByVal psFlag As Long) As Long
    Dim i As Long, player As Long, id As Long

    player = aFlagPos(psFlag).player

    'probably don't need to check mission list
    For i = 0 To 1
        Select Case i
        Case 0
            id = getRepairIdFromFlagSingleList(psFlag, player, apsStructLists(player))
            If id <> UDWORD_MAX Then getRepairIdFromFlag = id: Exit Function
        Case 1
            id = getRepairIdFromFlagSingleList(psFlag, player, mission_apsStructLists(player))
            If id <> UDWORD_MAX Then getRepairIdFromFlag = id: Exit Function
        End Select
    Next
    ASSERT_FAILURE "!""unable to find repair id for FLAG_POSITION""", "src/objmem.cpp:1092", "getRepairIdFromFlag", "getRepairIdFromFlag() failed"

    getRepairIdFromFlag = UDWORD_MAX
End Function

Public Sub objCount(droids As Long, structures As Long, features As Long)
    Dim i As Long, it As Long, psDroid As Long
    droids = 0
    structures = 0
    features = 0

    For i = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(i).head
        Do While it <> 0
            psDroid = lnObj(it)
            droids = droids + 1
            If DROID_isTransporter(psDroid) And aDroid(psDroid).psGroup <> 0 Then
                If aGroup(aDroid(psDroid).psGroup).psList.size <> 0 Then
                    droids = droids + aGroup(aDroid(psDroid).psGroup).psList.size
                End If
            End If
            it = lnNext(it)
        Loop

        structures = structures + apsStructLists(i).size
    Next

    features = features + apsFeatureList(0).size
End Sub
