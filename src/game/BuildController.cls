VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuildController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/hci/build.h, build.cpp: class BuildController : public BaseObjectsStatsController, public
' std::enable_shared_from_this<BuildController> (objects_stats_c's header: the three Implements are
' its bases; shared_from_this() is Me). Its typed methods: a STRUCTURE_STATS * is its index (-1 =
' nullptr), a DROID * its handle. The static members are build_c's BuildController_showFavorites
' and BuildController_highlightedBuilder; the widgets are build_c's.
Option Explicit

Implements BaseObjectsStatsController
Implements BaseStatsController
Implements BaseObjectsController

Private stats() As Long                 ' std::vector<STRUCTURE_STATS *>
Private stats_size As Long
Private builders() As Long              ' std::vector<DROID *>
Private builders_size As Long
Private queuedJumpToHighlightedStatsObject As Boolean
Private highlightedObjectStats As Long  ' BaseObjectsStatsController's (BASE_STATS *)

Private Sub Class_Initialize()
    queuedJumpToHighlightedStatsObject = True
End Sub

Public Function getStatsAt(ByVal statsIndex As Long) As Long
    If Not (statsIndex < stats_size) Then ASSERT_FAILURE "statsIndex < stats.size()", "src/hci/build.h:31", "getStatsAt", "Invalid stats index (%zu); max: (%zu)", statsIndex, stats_size: getStatsAt = -1: Exit Function
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
    updateBuildOptionsList
End Sub

Public Function shouldShowFavorites() As Boolean
    shouldShowFavorites = BuildController_showFavorites
End Function

Public Sub setShouldShowFavorite(ByVal value As Boolean)
    BuildController_showFavorites = value
    updateBuildOptionsList
End Sub

Public Function objectsSize() As Long
    objectsSize = builders_size
End Function

Public Function getObjectAt(ByVal index As Long) As Long
    If Not (index < builders_size) Then ASSERT_FAILURE "index < builders.size()", "src/hci/build.h:69", "getObjectAt", "Invalid object index (%zu); max: (%zu)", index, builders_size: getObjectAt = 0: Exit Function
    getObjectAt = builders(index)
End Function

Public Function findObject(ByVal iteration As WzFunc) As Boolean
    findObject = BaseObjectsController_findObject_vector(builders, builders_size, iteration)
End Function

Public Function getHighlightedObject() As Long
    getHighlightedObject = BuildController_highlightedBuilder
End Function

Public Function getQueuedJumpToHighlightedStatsObject() As Boolean
    getQueuedJumpToHighlightedStatsObject = queuedJumpToHighlightedStatsObject
End Function

Public Sub clearQueuedJumpToHighlightedStatsObject()
    queuedJumpToHighlightedStatsObject = False
End Sub

' ---- build.cpp

Public Sub updateData()
    updateBuildersList
    BaseObjectsController_updateHighlighted Me
    updateBuildOptionsList
End Sub

Private Sub updateBuildersList()
    Dim it As Long, droid_ As Long, i As Long, t As Long
    builders_size = 0
    Erase builders

    If Not (selectedPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/hci/build.cpp:46", "updateBuildersList", "selectedPlayer = %u", selectedPlayer: Exit Sub

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        droid_ = lnObj(it)
        If DROID_isConstructionDroid(droid_) And aDroid(droid_).died = 0 Then
            ReDim Preserve builders(0 To builders_size)
            builders(builders_size) = droid_
            builders_size = builders_size + 1
        End If
        it = lnNext(it)
    Loop

    ' std::reverse(builders.begin(), builders.end())
    For i = 0 To builders_size \ 2 - 1
        t = builders(i)
        builders(i) = builders(builders_size - 1 - i)
        builders(builders_size - 1 - i) = t
    Next
End Sub

Private Sub updateBuildOptionsList()
    Dim newBuildOptions() As Long, n As Long, i As Long
    n = fillStructureList(selectedPlayer, MAXSTRUCTURES - 1, shouldShowFavorites(), newBuildOptions)

    Erase stats
    stats_size = n
    If n > 0 Then
        ReDim stats(0 To n - 1)
        For i = 0 To n - 1
            stats(i) = newBuildOptions(i)
        Next
    End If
End Sub

Public Function getObjectStatsAt(ByVal objectIndex As Long) As Long
    Dim builder As Long, builderStats As Long, structure_ As Long
    builder = getObjectAt(objectIndex)
    If builder = 0 Then
        getObjectStatsAt = -1
        Exit Function
    End If

    If Not (aDroid(builder).droidType = DROID_CONSTRUCT Or aDroid(builder).droidType = DROID_CYBORG_CONSTRUCT) Then ' droidType(builder)
        getObjectStatsAt = -1
        Exit Function
    End If

    builderStats = -1
    If orderStateStatsLoc(builder, DORDER_BUILD, builderStats) Then ' Moving to build location?
        getObjectStatsAt = builderStats
        Exit Function
    End If

    structure_ = DroidGetBuildStructure(builder)
    If structure_ <> 0 Then ' Is building
        getObjectStatsAt = aStruct(structure_).pStructureType
        Exit Function
    End If

    If orderState(builder, DORDER_DEMOLISH) Then
        getObjectStatsAt = structGetDemolishStat()
        Exit Function
    End If

    getObjectStatsAt = -1
End Function


Public Sub startBuildPosition(ByVal buildOption As Long)
    Dim builder As Long
    builder = getHighlightedObject()
    If Not (builder <> 0) Then ASSERT_FAILURE "(builder) != nullptr", "src/hci/build.cpp:103", "startBuildPosition", "%s should not be nullptr", "builder": Exit Sub

    triggerEvent TRIGGER_MENU_BUILD_SELECTED

    If buildOption = structGetDemolishStat() Then
        objMode = IOBJ_DEMOLISHSEL
    Else
        objMode = IOBJ_BUILDSEL
        intStartConstructionPosition builder, buildOption
    End If

    intRemoveStats
    intMode = INT_OBJECT
End Sub

Public Sub toggleFavorites(ByVal buildOption As Long)
    ' bool &isFavorite = asStructureStats[buildOption->index].isFavorite;
    If shouldShowFavorites() Then
        asStructureStats(buildOption).isFavorite = False
    Else
        asStructureStats(buildOption).isFavorite = Not asStructureStats(buildOption).isFavorite
    End If
    updateBuildOptionsList
End Sub

Public Sub refresh()
    updateData

    If objectsSize() = 0 Then
        BaseObjectsController_closeInterface Me
    End If
End Sub

Public Sub clearData()
    builders_size = 0
    Erase builders
    setHighlightedObject 0, False
    stats_size = 0
    Erase stats
End Sub

Public Sub toggleBuilderSelection(ByVal droid_ As Long)
    Dim highlightedObject As Long
    If aDroid(droid_).selected <> 0 Then
        aDroid(droid_).selected = 0
    Else
        highlightedObject = getHighlightedObject()
        If highlightedObject <> 0 Then
            aDroid(highlightedObject).selected = 1
        End If
        BaseObjectsController_selectObject Me, droid_, False
    End If
    triggerEventSelected
End Sub

Public Sub setHighlightedObject(ByVal object_ As Long, ByVal jumpToHighlightedStatsObject As Boolean)
    Dim builder As Long
    If object_ = 0 Then
        BuildController_highlightedBuilder = 0
        Exit Sub
    End If

    builder = castDroid(object_)
    If Not (builder <> 0) Then ASSERT_FAILURE "(builder) != nullptr", "src/hci/build.cpp:178", "setHighlightedObject", "%s should not be nullptr", "builder": Exit Sub
    If Not (DROID_isConstructionDroid(builder)) Then ASSERT_FAILURE "builder->isConstructionDroid()", "src/hci/build.cpp:179", "setHighlightedObject", "Droid is not a construction droid": Exit Sub
    queuedJumpToHighlightedStatsObject = queuedJumpToHighlightedStatsObject Or jumpToHighlightedStatsObject
    BuildController_highlightedBuilder = builder
End Sub

Public Function showInterface() As Boolean
    Dim objectsForm As WIDGET
    updateData
    If builders_size = 0 Then
        showInterface = False
        Exit Function
    End If

    Set objectsForm = BuildObjectsForm_make(Me)
    psWScreen.psForm.attach objectsForm
    BaseStatsController_displayStatsForm Me
    triggerEvent TRIGGER_MENU_BUILD_UP
    showInterface = True
End Function

Public Function makeStatsForm() As WIDGET
    Set makeStatsForm = BuildStatsForm_make(Me)
End Function

' ---- the C++ bases (objects_stats_c's interfaces): BASE_STATS * handles

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
    BaseStatsController_getStatsAt = structureStatsBase(getStatsAt(index))
End Function

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
