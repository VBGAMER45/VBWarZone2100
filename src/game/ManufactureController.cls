VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ManufactureController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/hci/manufacture.h, manufacture.cpp: class ManufactureController : public
' BaseObjectsStatsController, public std::enable_shared_from_this<ManufactureController>
' (objects_stats_c's header). Its typed methods: a DROID_TEMPLATE * is a template pool handle (0 =
' nullptr), a STRUCTURE * its handle. The static member is manufacture_c's
' ManufactureController_highlightedFactory; the widgets are manufacture_c's.
Option Explicit

Implements BaseObjectsStatsController
Implements BaseStatsController
Implements BaseObjectsController

Private stats() As Long                 ' std::vector<DROID_TEMPLATE *>
Private stats_size As Long
Private factories() As Long             ' std::vector<STRUCTURE *>
Private factories_size As Long
Private queuedJumpToHighlightedStatsObject As Boolean
Private highlightedObjectStats As Long  ' BaseObjectsStatsController's (BASE_STATS *)

Private Sub Class_Initialize()
    queuedJumpToHighlightedStatsObject = True
End Sub

Public Function getStatsAt(ByVal statsIndex As Long) As Long
    If Not (statsIndex < stats_size) Then ASSERT_FAILURE "statsIndex < stats.size()", "src/hci/manufacture.h:31", "getStatsAt", "Invalid stats index (%zu); max: (%zu)", statsIndex, stats_size: getStatsAt = 0: Exit Function
    getStatsAt = stats(statsIndex)
End Function

Public Function statsSize() As Long
    statsSize = stats_size
End Function

Public Function shouldShowRedundantDesign() As Boolean
    shouldShowRedundantDesign = intGetShouldShowRedundantDesign()
End Function

Public Sub setShouldShowRedundantDesign(ByVal value As Boolean)
    intSetShouldShowRedundantDesign value
    updateManufactureOptionsList
End Sub

Public Function objectsSize() As Long
    objectsSize = factories_size
End Function

Public Function getObjectAt(ByVal index As Long) As Long
    If Not (index < factories_size) Then ASSERT_FAILURE "index < factories.size()", "src/hci/manufacture.h:58", "getObjectAt", "Invalid object index (%zu); max: (%zu)", index, factories_size: getObjectAt = 0: Exit Function
    getObjectAt = factories(index)
End Function

Public Function findObject(ByVal iteration As WzFunc) As Boolean
    findObject = BaseObjectsController_findObject_vector(factories, factories_size, iteration)
End Function

Public Function getHighlightedObject() As Long
    getHighlightedObject = ManufactureController_highlightedFactory
End Function

Public Function getQueuedJumpToHighlightedStatsObject() As Boolean
    getQueuedJumpToHighlightedStatsObject = queuedJumpToHighlightedStatsObject
End Function

Public Sub clearQueuedJumpToHighlightedStatsObject()
    queuedJumpToHighlightedStatsObject = False
End Sub

' ---- manufacture.cpp

Public Sub updateData()
    updateFactoriesList
    BaseObjectsController_updateHighlighted Me
    updateManufactureOptionsList
End Sub

Public Sub adjustFactoryProduction(ByVal manufactureOption As Long, ByVal add As Boolean)
    factoryProdAdjust getHighlightedObject(), manufactureOption, add
End Sub

Public Sub adjustFactoryLoop(ByVal add As Boolean)
    factoryLoopAdjust getHighlightedObject(), add
End Sub

Public Sub releaseFactoryProduction(ByVal structure_ As Long)
    releaseProduction structure_, ModeQueue
End Sub

Public Sub cancelFactoryProduction(ByVal structure_ As Long)
    If Not StructureIsManufacturingPending(structure_) Then
        Exit Sub
    End If

    If Not StructureIsOnHoldPending(structure_) Then
        holdProduction structure_, ModeQueue
        Exit Sub
    End If

    cancelProduction structure_, ModeQueue
    audio_PlayTrack ID_SOUND_WINDOWCLOSE
End Sub

Public Sub startDeliveryPointPosition()
    Dim factory_ As Long, psFlag As Long
    factory_ = getHighlightedObject()
    If Not (factory_ <> 0) Then ASSERT_FAILURE "(factory) != nullptr", "src/hci/manufacture.cpp:106", "startDeliveryPointPosition", "%s should not be nullptr", "factory": Exit Sub

    ' make sure that the factory isn't assigned to a commander
    assignFactoryCommandDroid factory_, 0
    psFlag = FindFactoryDelivery(factory_)
    If psFlag <> 0 Then
        startDeliveryPosition psFlag
    End If
End Sub

Private Sub updateFactoriesList()
    Dim intStrList As OBJ_LIST, it As Long, structure_ As Long, i As Long, j As Long, t As Long
    factories_size = 0
    Erase factories

    If interfaceStructList(intStrList) Then
        it = intStrList.head
        Do While it <> 0
            structure_ = lnObj(it)
            If aStruct(structure_).status = SS_BUILT And aStruct(structure_).died = 0 And STRUCTURE_isFactory(structure_) Then
                ReDim Preserve factories(0 To factories_size)
                factories(factories_size) = structure_
                factories_size = factories_size + 1
            End If
            it = lnNext(it)
        Loop
    End If

    ' std::sort(factories.begin(), factories.end(), compareFactories)
    For i = 1 To factories_size - 1
        t = factories(i)
        j = i - 1
        Do While j >= 0
            If Not compareFactories(t, factories(j)) Then Exit Do
            factories(j + 1) = factories(j)
            j = j - 1
        Loop
        factories(j + 1) = t
    Next
End Sub

Private Sub updateManufactureOptionsList()
    Dim factory_ As Long
    factory_ = getHighlightedObject()
    If factory_ <> 0 Then
        Erase stats
        stats_size = fillTemplateList(factory_, stats)
    Else
        stats_size = 0
        Erase stats
    End If
End Sub

Public Function getObjectStatsAt(ByVal objectIndex As Long) As Long
    Dim factory_ As Long
    factory_ = getFactoryOrNullptr(getObjectAt(objectIndex))
    If factory_ = 0 Then getObjectStatsAt = 0 Else getObjectStatsAt = aStruct(factory_).functionality.factory.psSubject
End Function

Public Sub refresh()
    updateData

    If objectsSize() = 0 Then
        BaseObjectsController_closeInterface Me
    End If
End Sub

Public Sub clearData()
    factories_size = 0
    Erase factories
    setHighlightedObject 0, False
    stats_size = 0
    Erase stats
End Sub

Public Sub setHighlightedObject(ByVal object_ As Long, ByVal jumpToHighlightedStatsObject As Boolean)
    Dim factory_ As Long, ok As Boolean
    If object_ = 0 Then
        ManufactureController_highlightedFactory = 0
        Exit Sub
    End If

    factory_ = castStructure(object_)
    If factory_ <> 0 Then ok = STRUCTURE_isFactory(factory_)
    If Not ok Then ASSERT_FAILURE "factory && factory->isFactory()", "src/hci/manufacture.cpp:200", "setHighlightedObject", "Invalid factory pointer": Exit Sub
    queuedJumpToHighlightedStatsObject = queuedJumpToHighlightedStatsObject Or jumpToHighlightedStatsObject
    ManufactureController_highlightedFactory = factory_
End Sub

Public Function showInterface() As Boolean
    Dim objectsForm As WIDGET
    updateData
    If factories_size = 0 Then
        showInterface = False
        Exit Function
    End If

    Set objectsForm = ManufactureObjectsForm_make(Me)
    psWScreen.psForm.attach objectsForm
    BaseStatsController_displayStatsForm Me
    triggerEvent TRIGGER_MENU_MANUFACTURE_UP
    showInterface = True
End Function

Public Function makeStatsForm() As WIDGET
    Set makeStatsForm = ManufactureStatsForm_make(Me)
End Function

' ---- the C++ bases (objects_stats_c's interfaces): BASE_STATS * handles (a template's: STAT_TEMPLATE + its handle)

Private Function templateBase(ByVal psTemplate As Long) As Long
    If psTemplate <> 0 Then templateBase = STAT_TEMPLATE + psTemplate Else templateBase = 0
End Function

Private Property Get BaseObjectsStatsController_highlightedObjectStats() As Long
    BaseObjectsStatsController_highlightedObjectStats = highlightedObjectStats
End Property

Private Property Let BaseObjectsStatsController_highlightedObjectStats(ByVal v As Long)
    highlightedObjectStats = v
End Property

Private Function BaseStatsController_statsSize() As Long
    BaseStatsController_statsSize = statsSize()
End Function

Private Function BaseStatsController_makeStatsForm() As WIDGET
    Set BaseStatsController_makeStatsForm = makeStatsForm()
End Function

Private Function BaseStatsController_getStatsAt(ByVal index As Long) As Long
    BaseStatsController_getStatsAt = templateBase(getStatsAt(index))
End Function

Private Function BaseObjectsController_objectsSize() As Long
    BaseObjectsController_objectsSize = objectsSize()
End Function

Private Function BaseObjectsController_getObjectAt(ByVal index As Long) As Long
    BaseObjectsController_getObjectAt = getObjectAt(index)
End Function

Private Function BaseObjectsController_getObjectStatsAt(ByVal index As Long) As Long
    BaseObjectsController_getObjectStatsAt = templateBase(getObjectStatsAt(index))
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
