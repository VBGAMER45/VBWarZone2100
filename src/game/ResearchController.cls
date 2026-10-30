VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResearchController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/hci/research.h, research.cpp: class ResearchController : public BaseObjectsStatsController,
' public std::enable_shared_from_this<ResearchController> (objects_stats_c's header). Its typed
' methods: a RESEARCH * is its index in asResearch (-1 = nullptr; a RESEARCH & too), a STRUCTURE *
' its handle. The static member is research_ui_c's ResearchController_highlightedFacility; the
' widgets are research_ui_c's.
Option Explicit

Implements BaseObjectsStatsController
Implements BaseStatsController
Implements BaseObjectsController

Private stats() As Long                 ' std::vector<RESEARCH *>
Private stats_size As Long
Private facilities() As Long            ' std::vector<STRUCTURE *>
Private facilities_size As Long
Private queuedJumpToHighlightedStatsObject As Boolean
Private highlightedObjectStats As Long  ' BaseObjectsStatsController's (BASE_STATS *)

Private Sub Class_Initialize()
    queuedJumpToHighlightedStatsObject = True
End Sub

Public Function getStatsAt(ByVal statsIndex As Long) As Long
    If Not (statsIndex < stats_size) Then ASSERT_FAILURE "statsIndex < stats.size()", "src/hci/research.h:31", "getStatsAt", "Invalid stats index (%zu); max: (%zu)", statsIndex, stats_size: getStatsAt = -1: Exit Function
    getStatsAt = stats(statsIndex)
End Function

Public Function statsSize() As Long
    statsSize = stats_size
End Function

Public Function objectsSize() As Long
    objectsSize = facilities_size
End Function

Public Function getObjectAt(ByVal index As Long) As Long
    If Not (index < facilities_size) Then ASSERT_FAILURE "index < facilities.size()", "src/hci/research.h:47", "getObjectAt", "Invalid object index (%zu); max: (%zu)", index, facilities_size: getObjectAt = 0: Exit Function
    getObjectAt = facilities(index)
End Function

Public Function findObject(ByVal iteration As WzFunc) As Boolean
    findObject = BaseObjectsController_findObject_vector(facilities, facilities_size, iteration)
End Function

Public Function getHighlightedObject() As Long
    getHighlightedObject = ResearchController_highlightedFacility
End Function

Public Function getQueuedJumpToHighlightedStatsObject() As Boolean
    getQueuedJumpToHighlightedStatsObject = queuedJumpToHighlightedStatsObject
End Function

Public Sub clearQueuedJumpToHighlightedStatsObject()
    queuedJumpToHighlightedStatsObject = False
End Sub

' ---- research.cpp

Public Sub updateData()
    updateFacilitiesList
    BaseObjectsController_updateHighlighted Me
    updateResearchOptionsList
End Sub

Private Sub updateFacilitiesList()
    Dim intStrList As OBJ_LIST, it As Long, psStruct As Long, i As Long, t As Long
    facilities_size = 0
    Erase facilities

    If interfaceStructList(intStrList) Then
        it = intStrList.head
        Do While it <> 0
            psStruct = lnObj(it)
            If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_RESEARCH And aStruct(psStruct).status = SS_BUILT And aStruct(psStruct).died = 0 Then
                ReDim Preserve facilities(0 To facilities_size)
                facilities(facilities_size) = psStruct
                facilities_size = facilities_size + 1
            End If
            it = lnNext(it)
        Loop
    End If

    ' std::reverse(facilities.begin(), facilities.end())
    For i = 0 To facilities_size \ 2 - 1
        t = facilities(i)
        facilities(i) = facilities(facilities_size - 1 - i)
        facilities(facilities_size - 1 - i) = t
    Next
End Sub

' nonstd::optional<size_t> getHighlightedFacilityIndex(): False = nullopt, else the index in index
Public Function getHighlightedFacilityIndex(index As Long) As Boolean
    Dim i As Long
    If getHighlightedObject() = 0 Then
        getHighlightedFacilityIndex = False
        Exit Function
    End If

    For i = 0 To facilities_size - 1
        If facilities(i) = getHighlightedObject() Then
            index = i
            getHighlightedFacilityIndex = True
            Exit Function
        End If
    Next
    getHighlightedFacilityIndex = False
End Function

Private Sub updateResearchOptionsList()
    Dim highlightedIndex As Long, research_ As Long, has_researchIndex As Boolean, researchIndex As Long
    Dim researchOptions() As Long, n As Long, i As Long, j As Long, t As Long
    stats_size = 0
    Erase stats

    research_ = -1
    If getHighlightedFacilityIndex(highlightedIndex) Then research_ = getObjectStatsAt(highlightedIndex)
    If research_ <> -1 Then
        has_researchIndex = True
        researchIndex = asResearch(research_).index
    End If

    n = fillResearchList(selectedPlayer, has_researchIndex, researchIndex, MAXRESEARCH, researchOptions)
    If n > 0 Then ReDim stats(0 To n - 1)
    For i = 0 To n - 1
        stats(stats_size) = researchOptions(i)
        stats_size = stats_size + 1
    Next

    ' std::sort(stats.begin(), stats.end(), [](const RESEARCH *a, const RESEARCH *b) {...})
    For i = 1 To stats_size - 1
        t = stats(i)
        j = i - 1
        Do While j >= 0
            If Not researchLess(t, stats(j)) Then Exit Do
            stats(j + 1) = stats(j)
            j = j - 1
        Loop
        stats(j + 1) = t
    Next
End Sub

' updateResearchOptionsList's sort lambda
Private Function researchLess(ByVal a As Long, ByVal b As Long) As Boolean
    Dim iconIdA As Long, iconIdB As Long
    iconIdA = mapIconToRID(asResearch(a).iconID)
    iconIdB = mapIconToRID(asResearch(b).iconID)

    If iconIdA <> iconIdB Then
        researchLess = iconIdA < iconIdB
        Exit Function
    End If

    researchLess = StrComp(asResearch(a).id, asResearch(b).id, vbBinaryCompare) < 0
End Function

Public Function getObjectStatsAt(ByVal objectIndex As Long) As Long
    Dim facility As Long, psSubjectPending As Long
    facility = getObjectAt(objectIndex)
    If facility = 0 Then
        getObjectStatsAt = -1
        Exit Function
    End If
    If aStruct(facility).pFunctionality = 0 Then
        getObjectStatsAt = -1
        Exit Function
    End If

    ' RESEARCH_FACILITY *psResearchFacility = &facility->pFunctionality->researchFacility; (never nullptr)
    psSubjectPending = aStruct(aStruct(facility).pFunctionality).functionality.researchFacility.psSubjectPending
    If psSubjectPending <> -1 Then
        If Not IsResearchCompleted(asPlayerResList(aStruct(facility).player, asResearch(psSubjectPending).index)) Then
            getObjectStatsAt = psSubjectPending
            Exit Function
        End If
    End If

    getObjectStatsAt = aStruct(aStruct(facility).pFunctionality).functionality.researchFacility.psSubject
End Function

Public Sub refresh()
    updateData

    If objectsSize() = 0 Then
        BaseObjectsController_closeInterface Me
    End If
End Sub

Public Sub clearData()
    facilities_size = 0
    Erase facilities
    setHighlightedObject 0, False
    stats_size = 0
    Erase stats
End Sub

' research: a RESEARCH & (its index)
Public Sub startResearch(ByVal research_ As Long)
    Dim facility As Long, psLab As Long
    triggerEvent TRIGGER_MENU_RESEARCH_SELECTED

    If Not (selectedPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/hci/research.cpp:144", "startResearch", "invalid player: %u", selectedPlayer: Exit Sub

    facility = getHighlightedObject()

    If Not (facility <> 0) Then ASSERT_FAILURE "(facility) != nullptr", "src/hci/research.cpp:148", "startResearch", "%s should not be nullptr", "facility": Exit Sub
    ' auto psResFacilty = &facility->pFunctionality->researchFacility;

    If aStruct(aStruct(facility).pFunctionality).functionality.researchFacility.psSubject <> -1 Then
        cancelResearch facility
    End If

    psLab = findResearchingFacilityByResearchIndex(selectedPlayer, asResearch(research_).ref - STAT_RESEARCH)
    If psLab <> 0 Then
        cancelResearch psLab 'Clear it out of this lab as we are now researching it in another.
    End If

    ' Say that we want to do research [sic].
    sendResearchStatus facility, asResearch(research_).ref - STAT_RESEARCH, selectedPlayer, True
    setStatusPendingStart_research aStruct(aStruct(facility).pFunctionality).functionality.researchFacility, research_ ' Tell UI that we are going to research.

    'stop the button from flashing once a topic has been chosen
    stopReticuleButtonFlash IDRET_RESEARCH
End Sub

Public Sub setHighlightedObject(ByVal object_ As Long, ByVal jumpToHighlightedStatsObject As Boolean)
    Dim facility As Long
    If object_ = 0 Then
        ResearchController_highlightedFacility = 0
        Exit Sub
    End If

    facility = castStructure(object_)
    If Not (facility <> 0) Then ASSERT_FAILURE "(facility) != nullptr", "src/hci/research.cpp:179", "setHighlightedObject", "%s should not be nullptr", "facility": Exit Sub
    If Not (asStructureStats(aStruct(facility).pStructureType).typ = REF_RESEARCH) Then ASSERT_FAILURE "facility->pStructureType->type == REF_RESEARCH", "src/hci/research.cpp:180", "setHighlightedObject", "Invalid facility pointer": Exit Sub
    queuedJumpToHighlightedStatsObject = queuedJumpToHighlightedStatsObject Or jumpToHighlightedStatsObject
    ResearchController_highlightedFacility = facility
End Sub

Public Function showInterface() As Boolean
    Dim objectsForm As WIDGET
    updateData
    If facilities_size = 0 Then
        showInterface = False
        Exit Function
    End If

    Set objectsForm = ResearchObjectsForm_make(Me)
    psWScreen.psForm.attach objectsForm
    BaseStatsController_displayStatsForm Me
    triggerEvent TRIGGER_MENU_RESEARCH_UP
    showInterface = True
End Function

Public Function makeStatsForm() As WIDGET
    Set makeStatsForm = ResearchStatsForm_make(Me)
End Function

Public Sub cancelResearch(ByVal facility As Long)
    research_c.cancelResearch facility, ModeQueue    ' ::cancelResearch
End Sub

Public Sub requestResearchCancellation(ByVal facility As Long)
    If Not structureIsResearchingPending(facility) Then
        Exit Sub
    End If

    If Not StructureIsOnHoldPending(facility) Then
        holdResearch facility, ModeQueue
        Exit Sub
    End If

    cancelResearch facility
    audio_PlayTrack ID_SOUND_WINDOWCLOSE
End Sub

' ---- the C++ bases (objects_stats_c's interfaces): BASE_STATS * handles (a research's: its ref)

Private Function researchBase(ByVal psResearch As Long) As Long
    If psResearch <> -1 Then researchBase = asResearch(psResearch).ref Else researchBase = 0
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
    BaseStatsController_getStatsAt = researchBase(getStatsAt(index))
End Function

Private Function BaseObjectsController_objectsSize() As Long
    BaseObjectsController_objectsSize = objectsSize()
End Function

Private Function BaseObjectsController_getObjectAt(ByVal index As Long) As Long
    BaseObjectsController_getObjectAt = getObjectAt(index)
End Function

Private Function BaseObjectsController_getObjectStatsAt(ByVal index As Long) As Long
    BaseObjectsController_getObjectStatsAt = researchBase(getObjectStatsAt(index))
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
