VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommanderController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/hci/commander.h, commander.cpp: class CommanderController : public BaseObjectsController,
' public std::enable_shared_from_this<CommanderController> (objects_stats_c's header). Its typed
' methods: a STRUCTURE_STATS * is its index (-1 = nullptr), a DROID * or STRUCTURE * its handle.
' The static member is commander_c's CommanderController_highlightedCommander; the widgets are
' commander_c's.
Option Explicit

Implements BaseObjectsController

Private commanders() As Long            ' std::vector<DROID *>
Private commanders_size As Long
Private queuedJumpToHighlightedStatsObject As Boolean

Private Sub Class_Initialize()
    queuedJumpToHighlightedStatsObject = True
End Sub

Public Function objectsSize() As Long
    objectsSize = commanders_size
End Function

Public Function getObjectAt(ByVal index As Long) As Long
    If Not (index < commanders_size) Then ASSERT_FAILURE "index < commanders.size()", "src/hci/commander.h:39", "getObjectAt", "Invalid object index (%zu); max: (%zu)", index, commanders_size: getObjectAt = 0: Exit Function
    getObjectAt = commanders(index)
End Function

Public Function findObject(ByVal iteration As WzFunc) As Boolean
    findObject = BaseObjectsController_findObject_vector(commanders, commanders_size, iteration)
End Function

Public Function getHighlightedObject() As Long
    getHighlightedObject = CommanderController_highlightedCommander
End Function

Public Function getQueuedJumpToHighlightedStatsObject() As Boolean
    getQueuedJumpToHighlightedStatsObject = queuedJumpToHighlightedStatsObject
End Function

Public Sub clearQueuedJumpToHighlightedStatsObject()
    queuedJumpToHighlightedStatsObject = False
End Sub

' ---- commander.cpp

Public Sub updateData()
    updateCommandersList
    BaseObjectsController_updateHighlighted Me
End Sub

Private Sub updateCommandersList()
    Dim it As Long, droid_ As Long, i As Long, j As Long, t As Long
    commanders_size = 0
    Erase commanders

    If Not (selectedPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/hci/commander.cpp:39", "updateCommandersList", "selectedPlayer = %u", selectedPlayer: Exit Sub

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        droid_ = lnObj(it)
        If aDroid(droid_).droidType = DROID_COMMAND And aDroid(droid_).died = 0 Then
            ReDim Preserve commanders(0 To commanders_size)
            commanders(commanders_size) = droid_
            commanders_size = commanders_size + 1
        End If
        it = lnNext(it)
    Loop

    ' Sort the list of commanders from lowest to highest id (using a lambda function defined within the sort call)
    For i = 1 To commanders_size - 1
        t = commanders(i)
        j = i - 1
        Do While j >= 0
            If Not ULT(aDroid(t).id, aDroid(commanders(j)).id) Then Exit Do
            commanders(j + 1) = commanders(j)
            j = j - 1
        Loop
        commanders(j + 1) = t
    Next
End Sub

Public Function getObjectStatsAt(ByVal objectIndex As Long) As Long
    Dim assignedFactory As Long
    assignedFactory = getAssignedFactoryAt(objectIndex)
    If assignedFactory = 0 Then getObjectStatsAt = -1 Else getObjectStatsAt = aStruct(assignedFactory).pStructureType
End Function

Public Function getAssignedFactoryAt(ByVal objectIndex As Long) As Long
    Dim droid_ As Long
    droid_ = getObjectAt(objectIndex)
    If droid_ = 0 Then getAssignedFactoryAt = 0 Else getAssignedFactoryAt = droidGetCommandFactory(droid_)
End Function

Public Sub refresh()
    updateData

    If objectsSize() = 0 Then
        BaseObjectsController_closeInterface Me
    ' If order form is visible, refresh it.
    ElseIf Not widgGetFromID(psWScreen, IDORDER_FORM) Is Nothing Then
        displayOrderForm
    End If
End Sub

Public Sub clearData()
    commanders_size = 0
    Erase commanders
    setHighlightedObject 0, False
End Sub

Public Sub setHighlightedObject(ByVal object_ As Long, ByVal jumpToHighlightedStatsObject As Boolean)
    Dim commander As Long
    If object_ = 0 Then
        CommanderController_highlightedCommander = 0
        Exit Sub
    End If

    commander = castDroid(object_)
    If Not (commander <> 0) Then ASSERT_FAILURE "(commander) != nullptr", "src/hci/commander.cpp:97", "setHighlightedObject", "%s should not be nullptr", "commander": Exit Sub
    If Not (aDroid(commander).droidType = DROID_COMMAND) Then ASSERT_FAILURE "commander->droidType == DROID_COMMAND", "src/hci/commander.cpp:98", "setHighlightedObject", "Droid is not a commander": Exit Sub
    queuedJumpToHighlightedStatsObject = queuedJumpToHighlightedStatsObject Or jumpToHighlightedStatsObject
    CommanderController_highlightedCommander = commander
End Sub

Public Function showInterface() As Boolean
    Dim objectsForm As WIDGET
    updateData
    If commanders_size = 0 Then
        showInterface = False
        Exit Function
    End If

    Set objectsForm = CommanderObjectsForm_make(Me)
    psWScreen.psForm.attach objectsForm
    displayOrderForm
    showInterface = True
End Function

Public Sub displayOrderForm()
    Dim psWeakControllerRef As CommanderController
    Set psWeakControllerRef = Me         ' std::weak_ptr<CommanderController>(shared_from_this()): see commander_c's lambda
    widgScheduleTask WzFunc_new(FN_CommanderController_displayOrderForm_task, psWeakControllerRef)
End Sub

' ---- the C++ base (objects_stats_c's interface): BASE_STATS * handles

Private Function BaseObjectsController_objectsSize() As Long
    BaseObjectsController_objectsSize = objectsSize()
End Function

Private Function BaseObjectsController_getObjectAt(ByVal index As Long) As Long
    BaseObjectsController_getObjectAt = getObjectAt(index)
End Function

Private Function BaseObjectsController_getObjectStatsAt(ByVal index As Long) As Long
    BaseObjectsController_getObjectStatsAt = structureStatsBase(getObjectStatsAt(index))
End Function

Private Function BaseObjectsController_findObject(ByVal iteration As WzFunc) As Boolean
    BaseObjectsController_findObject = findObject(iteration)
End Function

Private Sub BaseObjectsController_refresh()
    refresh
End Sub

Private Function BaseObjectsController_showInterface() As Boolean
    BaseObjectsController_showInterface = showInterface()
End Function

Private Sub BaseObjectsController_prepareToClose()
    BaseObjectsController_def_prepareToClose Me
End Sub

Private Sub BaseObjectsController_clearData()
    clearData
End Sub

Private Function BaseObjectsController_getHighlightedObject() As Long
    BaseObjectsController_getHighlightedObject = getHighlightedObject()
End Function

Private Sub BaseObjectsController_setHighlightedObject(ByVal object_ As Long, ByVal jumpToHighlightedStatsObject As Boolean)
    setHighlightedObject object_, jumpToHighlightedStatsObject
End Sub

Private Function BaseObjectsController_getQueuedJumpToHighlightedStatsObject() As Boolean
    BaseObjectsController_getQueuedJumpToHighlightedStatsObject = getQueuedJumpToHighlightedStatsObject()
End Function

Private Sub BaseObjectsController_clearQueuedJumpToHighlightedStatsObject()
    clearQueuedJumpToHighlightedStatsObject
End Sub
