Attribute VB_Name = "research_ui_c"
' src/hci/research.cpp, research.h: the research panel's widgets, getResearchObjectImage,
' AllyResearchsIcons' display function and ResearchController's static member (the controller
' itself is the class ResearchController; AllyResearchsIcons is a class too). research_c is the
' simulation's research.cpp.
' (VB) widget class ResearchObjectButton : ObjectButton
' (VB) widget class ResearchStatsButton : StatsButton
' (VB) widget class ResearchOptionButton : StatsFormButton
' (VB) widget class ResearchObjectsForm : ObjectsForm
' (VB) widget class ResearchStatsForm : ObjectStatsForm
Option Explicit

Public ResearchController_highlightedFacility As Long ' static STRUCTURE *ResearchController::highlightedFacility = nullptr

Public Const FN_AllyResearchsIcons_displayAllyIcon As Long = 8300 ' static void AllyResearchsIcons::displayAllyIcon
Public Const FN_ResearchOptionButton_removeStats As Long = 8301   ' clickPrimary's widgScheduleTask lambda

' (VB) std::min of two ints
Private Function minL(ByVal a As Long, ByVal b As Long) As Long
    If b < a Then minL = b Else minL = a
End Function

' (VB) a RESEARCH * (its index, -1 = nullptr) as a BASE_STATS * handle (0 = nullptr)
Private Function researchBase(ByVal psResearch As Long) As Long
    If psResearch <> -1 Then researchBase = asResearch(psResearch).ref Else researchBase = 0
End Function

' MARK: - AllyResearchsIcons

Private Sub displayAllyIcon(ByVal widget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim label As WIDGET, x As Long, y As Long
    Set label = widget                  ' auto *label = (W_LABEL *)widget;
    x = label.x() + xOffset
    y = label.y() + yOffset

    iV_DrawImageTc_ids IntImages, IMAGE_ALLY_RESEARCH, IMAGE_ALLY_RESEARCH_TC, x, y, pal_GetTeamColour(label.UserData)
End Sub

' MARK: - ResearchObjectButton

Public Sub ResearchObjectButton_construct(ByVal self As WIDGET)
    ObjectButton_construct self
    self.kind = K_ResearchObjectButton
    Set ResearchObjectButton_(self).allyResearchIcons = New AllyResearchsIcons
End Sub

Public Function ResearchObjectButton_make(ByVal controller As ResearchController, ByVal objectIndex As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_ResearchObjectButton)
    ResearchObjectButton_construct widget
    ResearchObjectButton_initialize widget
    Set ResearchObjectButton_(widget).controller = controller
    ObjectButton_(widget).objectIndex = objectIndex
    Set ResearchObjectButton_make = widget
End Function

Public Sub ResearchObjectButton_initialize(ByVal self As WIDGET)
    If bMultiPlayer Then
        ResearchObjectButton_(self).allyResearchIcons.initialize self
        ResearchObjectButton_(self).allyResearchIcons.setTop 10
    End If
End Sub

Public Sub ResearchObjectButton_jump(ByVal self As WIDGET)
    If Not offWorldKeepLists Then
        ObjectButton_jump self
    End If
End Sub

Public Sub ResearchObjectButton_clickPrimary(ByVal self As WIDGET)
    Dim controller As ResearchController
    Set controller = ResearchObjectButton_(self).controller
    BaseObjectsController_clearStructureSelection controller
    BaseObjectsController_selectObject controller, controller.getObjectAt(ObjectButton_(self).objectIndex), False
    vt_ObjectButton_jump self
    BaseStatsController_scheduleDisplayStatsForm controller
End Sub

Public Sub ResearchObjectButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim facility As Long
    vt_DynamicIntFancyButton_updateLayout self
    facility = ResearchObjectButton_(self).controller.getObjectAt(ObjectButton_(self).objectIndex)
    If Not (facility <> 0) Then ASSERT_FAILURE "(facility) != nullptr", "src/hci/research.cpp:292", "display", "%s should not be nullptr", "facility": Exit Sub
    If isDead(facility) Then
        ASSERT_FAILURE "!isDead(facility)", "src/hci/research.cpp:295", "display", "Facility is dead"
        ' ensure the backing information is refreshed before the next draw
        intRefreshScreen
        Exit Sub
    End If
    IntFancyButton_displayIMD self, AtlasImage_(), ImdObject_Structure(facility), xOffset, yOffset
    IntFancyButton_displayIfHighlight self, xOffset, yOffset
End Sub

Public Sub ResearchObjectButton_updateLayout(ByVal self As WIDGET)
    DynamicIntFancyButton_updateLayout self

    If bMultiPlayer Then
        ResearchObjectButton_updateAllyStatus self
    End If
End Sub

Public Sub ResearchObjectButton_updateAllyStatus(ByVal self As WIDGET)
    Dim research_ As Long, allyResearches() As AllyResearch, n As Long
    ResearchObjectButton_(self).allyResearchIcons.hide

    research_ = ResearchObjectButton_(self).controller.getObjectStatsAt(ObjectButton_(self).objectIndex)
    If research_ = -1 Then
        Exit Sub
    End If

    n = listAllyResearch(asResearch(research_).ref, allyResearches)
    If n = 0 Then
        Exit Sub
    End If

    ResearchObjectButton_(self).allyResearchIcons.update allyResearches, n
End Sub

Public Function ResearchObjectButton_getTip(ByVal self As WIDGET) As String
    Dim facility As Long
    facility = ResearchObjectButton_(self).controller.getObjectAt(ObjectButton_(self).objectIndex)
    If Not (facility <> 0) Then ASSERT_FAILURE "(facility) != nullptr", "src/hci/research.cpp:336", "getTip", "%s should not be nullptr", "facility": ResearchObjectButton_getTip = "": Exit Function
    ResearchObjectButton_getTip = BASE_STATS_name(structureStatsBase(aStruct(facility).pStructureType)) ' getLocalizedStatsName(facility->pStructureType)
End Function

Public Function ResearchObjectButton_getController(ByVal self As WIDGET) As Object
    Set ResearchObjectButton_getController = ResearchObjectButton_(self).controller
End Function

' MARK: - ResearchStatsButton

Public Sub ResearchStatsButton_construct(ByVal self As WIDGET)
    StatsButton_construct self
    self.kind = K_ResearchStatsButton
End Sub

Public Function ResearchStatsButton_make(ByVal controller As ResearchController, ByVal objectIndex As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_ResearchStatsButton)
    ResearchStatsButton_construct widget
    Set ResearchStatsButton_(widget).controller = controller
    ResearchStatsButton_(widget).objectIndex = objectIndex
    ResearchStatsButton_initialize widget
    Set ResearchStatsButton_make = widget
End Function

Public Sub ResearchStatsButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim facility As Long, research_ As Long, researchPending As Boolean, objectImage As ImdObject
    vt_DynamicIntFancyButton_updateLayout self

    facility = ResearchStatsButton_(self).controller.getObjectAt(ResearchStatsButton_(self).objectIndex)
    research_ = ResearchStatsButton_(self).controller.getObjectStatsAt(ResearchStatsButton_(self).objectIndex)
    If facility <> 0 Then researchPending = structureIsResearchingPending(facility)
    If researchPending And research_ <> -1 Then objectImage = getResearchObjectImage(research_) Else objectImage = ImdObject_Component(0)

    IntFancyButton_displayIMD self, AtlasImage_(), objectImage, xOffset, yOffset

    If researchPending And StructureIsOnHoldPending(facility) Then
        If ((realTime \ 250) Mod 2) = 0 Then
            iV_DrawImage IntImages, IMAGE_BUT0_DOWN, xOffset + self.x(), yOffset + self.y()
        Else
            iV_DrawImage IntImages, IMAGE_BUT_HILITE, xOffset + self.x(), yOffset + self.y()
        End If
    Else
        IntFancyButton_displayIfHighlight self, xOffset, yOffset
    End If
End Sub

Public Sub ResearchStatsButton_updateLayout(ByVal self As WIDGET)
    DynamicIntFancyButton_updateLayout self
    ResearchStatsButton_updateProgressBar self
End Sub

' RESEARCH *getStats() override: as the BASE_STATS * of the virtual
Public Function ResearchStatsButton_getStats(ByVal self As WIDGET) As Long
    ResearchStatsButton_getStats = researchBase(ResearchStatsButton_(self).controller.getObjectStatsAt(ResearchStatsButton_(self).objectIndex))
End Function

Public Sub ResearchStatsButton_initialize(ByVal self As WIDGET)
    StatsButton_addProgressBar self
End Sub

Public Sub ResearchStatsButton_updateProgressBar(ByVal self As WIDGET)
    Dim progressBar As WIDGET, facility As Long, research_ As Long, psSubject As Long, currentPoints As Long
    Dim researchRate As Long, neededPower As Long, powerToBuild As Long
    Set progressBar = StatsButton_(self).progressBar
    progressBar.hide
    facility = ResearchStatsButton_(self).controller.getObjectAt(ResearchStatsButton_(self).objectIndex)
    If Not (facility <> 0) Then ASSERT_FAILURE "(facility) != nullptr", "src/hci/research.cpp:410", "updateProgressBar", "%s should not be nullptr", "facility": Exit Sub

    research_ = StructureGetResearch(facility)
    If research_ = 0 Then
        Exit Sub
    End If
    psSubject = aStruct(research_).functionality.researchFacility.psSubject
    If psSubject = -1 Then
        Exit Sub
    End If

    If Not (selectedPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/hci/research.cpp:422", "updateProgressBar", "invalid player: %u", selectedPlayer: Exit Sub
    ' auto& playerResList = asPlayerResList[selectedPlayer];
    If Not (asResearch(psSubject).index < asPlayerResList_size(selectedPlayer)) Then ASSERT_FAILURE "research->psSubject->index < playerResList.size()", "src/hci/research.cpp:424", "updateProgressBar", "Invalid index": Exit Sub
    currentPoints = asPlayerResList(selectedPlayer, asResearch(psSubject).index).currentPoints
    If currentPoints <> 0 Then
        If aStruct(research_).functionality.researchFacility.timeStartHold = 0 Then researchRate = getBuildingResearchPoints(facility) Else researchRate = 0
        formatTime progressBar, currentPoints, asResearch(psSubject).researchPoints, researchRate, "Research Progress"
    Else
        ' Not yet started production.
        neededPower = checkPowerRequest(facility)
        powerToBuild = asResearch(psSubject).researchPower  ' research->psSubject != nullptr ? ... : 0
        formatPower progressBar, neededPower, powerToBuild
    End If
End Sub

Public Function ResearchStatsButton_isHighlighted(ByVal self As WIDGET) As Boolean
    Dim controller As ResearchController, facility As Long
    Set controller = ResearchStatsButton_(self).controller
    facility = controller.getObjectAt(ResearchStatsButton_(self).objectIndex)
    If facility <> 0 Then ResearchStatsButton_isHighlighted = (aStruct(facility).selected <> 0 Or facility = controller.getHighlightedObject())
End Function

Public Sub ResearchStatsButton_clickPrimary(ByVal self As WIDGET)
    Dim controller As ResearchController, facility As Long
    Set controller = ResearchStatsButton_(self).controller
    facility = controller.getObjectAt(ResearchStatsButton_(self).objectIndex)
    If Not (facility <> 0) Then ASSERT_FAILURE "(facility) != nullptr", "src/hci/research.cpp:449", "clickPrimary", "%s should not be nullptr", "facility": Exit Sub

    'might need to cancel the hold on research facility
    releaseResearch facility, ModeQueue
    BaseObjectsController_clearStructureSelection controller
    BaseObjectsController_selectObject controller, facility, True
    BaseStatsController_scheduleDisplayStatsForm controller
    controller.refresh
End Sub

Public Sub ResearchStatsButton_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean)
    Dim controller As ResearchController, facility As Long
    Set controller = ResearchStatsButton_(self).controller
    facility = controller.getObjectAt(ResearchStatsButton_(self).objectIndex)
    If Not (facility <> 0) Then ASSERT_FAILURE "(facility) != nullptr", "src/hci/research.cpp:462", "clickSecondary", "%s should not be nullptr", "facility": Exit Sub
    BaseObjectsController_clearStructureSelection controller
    controller.requestResearchCancellation facility
    controller.setHighlightedObject facility, True
    BaseStatsController_scheduleDisplayStatsForm controller
    controller.refresh
End Sub

' MARK: - ResearchOptionButton

Public Sub ResearchOptionButton_construct(ByVal self As WIDGET)
    StatsFormButton_construct self
    self.kind = K_ResearchOptionButton
    Set ResearchOptionButton_(self).allyResearchIcons = New AllyResearchsIcons
End Sub

Public Function ResearchOptionButton_make(ByVal controller As ResearchController, ByVal researchOptionIndex As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_ResearchOptionButton)
    ResearchOptionButton_construct widget
    Set ResearchOptionButton_(widget).controller = controller
    ResearchOptionButton_(widget).researchOptionIndex = researchOptionIndex
    ResearchOptionButton_initialize widget
    Set ResearchOptionButton_make = widget
End Function

' RESEARCH *getStats() override: as the BASE_STATS * of the virtual
Public Function ResearchOptionButton_getStats(ByVal self As WIDGET) As Long
    ResearchOptionButton_getStats = researchBase(ResearchOptionButton_(self).controller.getStatsAt(ResearchOptionButton_(self).researchOptionIndex))
End Function

' (VB) getStats() as the RESEARCH * (its index, -1 = nullptr) the C++ override returns
Private Function ResearchOptionButton_research(ByVal self As WIDGET) As Long
    ResearchOptionButton_research = ResearchOptionButton_(self).controller.getStatsAt(ResearchOptionButton_(self).researchOptionIndex)
End Function

Public Sub ResearchOptionButton_initialize(ByVal self As WIDGET)
    StatsFormButton_addCostBar self

    If bMultiPlayer Then
        StatsButton_addProgressBar self
        ResearchOptionButton_(self).allyResearchIcons.initialize self
        StatsButton_(self).progressBar.move STAT_TIMEBARX, STAT_TIMEBARY
        StatsButton_(self).progressBar.setTip "Ally progress"
    End If
End Sub

Public Sub ResearchOptionButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim research_ As Long, researchIcon As AtlasImage
    vt_DynamicIntFancyButton_updateLayout self

    research_ = ResearchOptionButton_research(self)
    If Not (research_ <> -1) Then ASSERT_FAILURE "(research) != nullptr", "src/hci/research.cpp:516", "display", "%s should not be nullptr", "research": Exit Sub

    If asResearch(research_).iconID <> NO_RESEARCH_ICON Then researchIcon = AtlasImage_(IntImages, asResearch(research_).iconID) Else researchIcon = AtlasImage_()
    IntFancyButton_displayIMD self, researchIcon, getResearchObjectImage(research_), xOffset, yOffset

    If asResearch(research_).subGroup <> NO_RESEARCH_ICON Then
        iV_DrawImage IntImages, asResearch(research_).subGroup, xOffset + self.x() + STAT_BUTWIDTH - 16, yOffset + self.y() + 3
    End If

    IntFancyButton_displayIfHighlight self, xOffset, yOffset
End Sub

Public Function ResearchOptionButton_isHighlighted(ByVal self As WIDGET) As Boolean
    ResearchOptionButton_isHighlighted = BaseObjectsStatsController_isHighlightedObjectStats(ResearchOptionButton_(self).controller, ResearchOptionButton_(self).researchOptionIndex)
End Function

Public Sub ResearchOptionButton_updateLayout(ByVal self As WIDGET)
    Dim research_ As Long
    DynamicIntFancyButton_updateLayout self
    research_ = ResearchOptionButton_research(self)

    ResearchOptionButton_updateCostBar self, research_

    If bMultiPlayer Then
        ResearchOptionButton_updateAllyStatus self, research_
    End If

    If self.isMouseOverWidget() Then
        intSetShadowPower vt_StatsFormButton_getCost(self)
    End If
End Sub

Public Sub ResearchOptionButton_updateCostBar(ByVal self As WIDGET, ByVal stat As Long)
    W_BARGRAPH_(StatsFormButton_(self).costBar).majorSize = minL(100, vt_StatsFormButton_getCost(self) \ POWERPOINTS_DROIDDIV)
End Sub

Public Sub ResearchOptionButton_updateAllyStatus(ByVal self As WIDGET, ByVal research_ As Long)
    Dim costBar As WIDGET, progressBar As WIDGET, allyResearches() As AllyResearch, n As Long
    Set costBar = StatsFormButton_(self).costBar
    Set progressBar = StatsButton_(self).progressBar
    costBar.move STAT_TIMEBARX, STAT_TIMEBARY
    progressBar.hide
    ResearchOptionButton_(self).allyResearchIcons.hide

    If research_ = -1 Then
        Exit Sub
    End If

    n = listAllyResearch(asResearch(research_).ref, allyResearches)
    If n = 0 Then
        Exit Sub
    End If

    intDisplayUpdateAllyBar progressBar, asResearch(research_), allyResearches, n

    ResearchOptionButton_(self).allyResearchIcons.update allyResearches, n

    costBar.move STAT_TIMEBARX, STAT_TIMEBARY - STAT_PROGBARHEIGHT - 2
    progressBar.show
End Sub

Public Function ResearchOptionButton_getCost(ByVal self As WIDGET) As Long
    Dim research_ As Long
    research_ = ResearchOptionButton_research(self)
    If research_ <> -1 Then ResearchOptionButton_getCost = asResearch(research_).researchPower Else ResearchOptionButton_getCost = 0
End Function

Public Sub ResearchOptionButton_clickPrimary(ByVal self As WIDGET)
    Dim controller As ResearchController, clickedStats As Long
    Set controller = ResearchOptionButton_(self).controller
    clickedStats = controller.getStatsAt(ResearchOptionButton_(self).researchOptionIndex)
    If Not (clickedStats <> -1) Then ASSERT_FAILURE "(clickedStats) != nullptr", "src/hci/research.cpp:591", "clickPrimary", "%s should not be nullptr", "clickedStats": Exit Sub
    If BaseObjectsStatsController_isHighlightedObjectStats(controller, ResearchOptionButton_(self).researchOptionIndex) Then
        controller.cancelResearch controller.getHighlightedObject()
    Else
        controller.startResearch clickedStats
    End If
    widgScheduleTask WzFunc_new(FN_ResearchOptionButton_removeStats)
End Sub

Public Function ResearchOptionButton_getTip(ByVal self As WIDGET) As String
    Dim costString As String, stats As Long, tipString As String, facility As Long
    Dim currentPoints As Long, totalPoints As Long, researchRate As Long, timeToBuild As Long, timeText As String
    costString = va("Cost: %u", vt_StatsFormButton_getCost(self)) ' WzString::format(_("Cost: %u"), getCost())
    stats = ResearchOptionButton_research(self)
    If stats = -1 Then tipString = "" Else tipString = BASE_STATS_name(researchBase(stats)) ' getLocalizedStatsName(stats)
    If stats <> -1 Then
        If Len(asResearch(stats).category) <> 0 Then
            tipString = tipString & va(vbLf & "(%s %u/%u)", asResearch(stats).category, asResearch(stats).categoryProgress, asResearch(stats).categoryMax)
        End If
    End If
    tipString = tipString & vbLf
    tipString = tipString & costString

    facility = ResearchOptionButton_(self).controller.getHighlightedObject()
    If stats <> -1 And facility <> 0 And selectedPlayer < MAX_PLAYERS Then
        ' auto& playerResList = asPlayerResList[selectedPlayer];
        If asResearch(stats).index < asPlayerResList_size(selectedPlayer) Then
            currentPoints = asPlayerResList(selectedPlayer, asResearch(stats).index).currentPoints
            totalPoints = asResearch(stats).researchPoints
            researchRate = getBuildingResearchPoints(facility)
            If researchRate > 0 Then
                timeToBuild = (totalPoints - currentPoints) \ researchRate
                timeText = va("%u:%02u", timeToBuild \ 60, timeToBuild Mod 60)
                tipString = tipString & " ("
                tipString = tipString & timeText
                tipString = tipString & ")"
            End If
        End If
    End If

    ResearchOptionButton_getTip = tipString
End Function

' MARK: - ResearchObjectsForm

Public Sub ResearchObjectsForm_construct(ByVal self As WIDGET)
    ObjectsForm_construct self
    self.kind = K_ResearchObjectsForm
End Sub

Public Function ResearchObjectsForm_make(ByVal controller As ResearchController) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_ResearchObjectsForm)
    ResearchObjectsForm_construct widget
    Set ResearchObjectsForm_(widget).controller = controller
    ObjectsForm_initialize widget
    Set ResearchObjectsForm_make = widget
End Function

Public Function ResearchObjectsForm_makeStatsButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set ResearchObjectsForm_makeStatsButton = ResearchStatsButton_make(ResearchObjectsForm_(self).controller, buttonIndex)
End Function

Public Function ResearchObjectsForm_makeObjectButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set ResearchObjectsForm_makeObjectButton = ResearchObjectButton_make(ResearchObjectsForm_(self).controller, buttonIndex)
End Function

Public Function ResearchObjectsForm_getController(ByVal self As WIDGET) As Object
    Set ResearchObjectsForm_getController = ResearchObjectsForm_(self).controller
End Function

' MARK: - ResearchStatsForm

Public Sub ResearchStatsForm_construct(ByVal self As WIDGET)
    ObjectStatsForm_construct self
    self.kind = K_ResearchStatsForm
End Sub

Public Function ResearchStatsForm_make(ByVal controller As ResearchController) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_ResearchStatsForm)
    ResearchStatsForm_construct widget
    Set ResearchStatsForm_(widget).controller = controller
    vt_StatsForm_initialize widget
    Set ResearchStatsForm_make = widget
End Function

Public Function ResearchStatsForm_makeOptionButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set ResearchStatsForm_makeOptionButton = ResearchOptionButton_make(ResearchStatsForm_(self).controller, buttonIndex)
End Function

Public Function ResearchStatsForm_getController(ByVal self As WIDGET) As Object
    Set ResearchStatsForm_getController = ResearchStatsForm_(self).controller
End Function

' research: a RESEARCH * (its index); the ImdObject's stats are BASE_STATS * handles
Public Function getResearchObjectImage(ByVal research_ As Long) As ImdObject
    Dim psResGraphic As Long, compID As Long, none As ImdObject
    psResGraphic = asResearch(research_).psStat

    If psResGraphic = 0 Then
        ' no Stat for this research topic so just use the graphic provided
        ' if Object != NULL the there must be a IMD so set the object to
        ' equal the Research stat
        If asResearch(research_).pIMD <> 0 Then
            getResearchObjectImage = ImdObject_Research(researchBase(research_))
            Exit Function
        End If

        getResearchObjectImage = none   ' ImdObject()
        Exit Function
    End If

    ' we have a Stat associated with this research topic
    If StatIsStructure(psResGraphic) Then
        ' overwrite the Object pointer
        getResearchObjectImage = ImdObject_StructureStat(psResGraphic)
        Exit Function
    End If

    compID = StatIsComponent(psResGraphic)
    If compID <> COMP_NUMCOMPONENTS Then
        getResearchObjectImage = ImdObject_Component(psResGraphic)
        Exit Function
    End If

    ASSERT_FAILURE "false", "src/hci/research.cpp:785", "getResearchObjectImage", "Invalid Stat for research button"
    getResearchObjectImage = ImdObject_Research(0)
End Function

' MARK: - the lambdas

Public Function research_ui_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Select Case f.id
    Case FN_AllyResearchsIcons_displayAllyIcon
        displayAllyIcon args(0), args(1), args(2)
    Case FN_ResearchOptionButton_removeStats
        intRemoveStats
    End Select
End Function
