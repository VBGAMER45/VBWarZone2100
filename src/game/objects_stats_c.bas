Attribute VB_Name = "objects_stats_c"
' src/hci/objects_stats.cpp, objects_stats.h: the bases of the object and stats panels (build,
' manufacture, research, commander). The controllers are not widgets: BaseObjectsController,
' BaseStatsController and BaseObjectsStatsController are VB interface classes of their virtual
' methods (and of the field highlightedObjectStats), and a concrete controller class Implements
' every C++ base it derives from (C++'s multiple inheritance written out: BuildController
' Implements BaseObjectsStatsController, BaseStatsController and BaseObjectsController, each
' Implements procedure forwarding to the class's own public method). Their non-virtual methods are
' this module's BaseObjectsController_<method>(controller, ...), BaseStatsController_<method> and
' BaseObjectsStatsController_<method>. A BASE_OBJECT * is its pool handle and a BASE_STATS * its
' stats handle (0 = nullptr) in the interfaces; a controller's own typed methods return its typed
' pointers (STRUCTURE_STATS *, RESEARCH *: an index, -1 = nullptr; DROID_TEMPLATE *: a template
' pool handle, 0 = nullptr). A std::shared_ptr<XController> is the class object, a
' std::weak_ptr<XController>'s lock the object itself (see BaseStatsController_scheduleDisplayStatsForm).
' (VB) widget class DynamicIntFancyButton : IntFancyButton
' (VB) virtual DynamicIntFancyButton.isHighlighted
' (VB) virtual DynamicIntFancyButton.updateLayout
' (VB) virtual DynamicIntFancyButton.clickPrimary
' (VB) virtual DynamicIntFancyButton.clickSecondary
' (VB) widget class StatsButton : DynamicIntFancyButton
' (VB) virtual StatsButton.getStats
' (VB) widget class ObjectButton : DynamicIntFancyButton
' (VB) virtual ObjectButton.getController
' (VB) virtual ObjectButton.jump
' (VB) widget class StatsFormButton : StatsButton
' (VB) virtual StatsFormButton.getCost
' (VB) widget class ObjectsForm : IntFormAnimated
' (VB) virtual ObjectsForm.makeStatsButton
' (VB) virtual ObjectsForm.makeObjectButton
' (VB) virtual ObjectsForm.getController
' (VB) widget class StatsForm : IntFormAnimated
' (VB) virtual StatsForm.initialize
' (VB) virtual StatsForm.updateLayout
' (VB) virtual StatsForm.makeOptionButton
' (VB) virtual StatsForm.getController
' (VB) widget class ObjectStatsForm : StatsForm
Option Explicit

Public Const STAT_GAP As Long = 2
Public Const STAT_BUTWIDTH As Long = 60
Public Const STAT_BUTHEIGHT As Long = 46

Public Const FN_closeInterface_task As Long = 8000                  ' closeInterface's widgScheduleTask lambda
Public Const FN_closeInterfaceNoAnim_task As Long = 8001            ' closeInterfaceNoAnim's
Public Const FN_findAnySelectedObject As Long = 8002                ' updateHighlighted's lambdas (cap 0: the controller)
Public Const FN_findHighlightedObject As Long = 8003                '   (cap 1: highlighted)
Public Const FN_scheduleDisplayStatsForm_task As Long = 8004        ' (cap 0: the controller)
Public Const FN_ObjectsForm_calcLayout As Long = 8005
Public Const FN_ObjectsForm_closeCalcLayout As Long = 8006
Public Const FN_ObjectsForm_closeClicked As Long = 8007
Public Const FN_ObjectsForm_closeTask As Long = 8008
Public Const FN_ObjectsForm_tabListCalcLayout As Long = 8009
Public Const FN_StatsForm_calcLayout As Long = 8010
Public Const FN_StatsForm_closeClicked As Long = 8011
Public Const FN_StatsForm_closeTask As Long = 8012

' MARK: - BaseObjectsController (c: any controller; VB casts it to the interface)

' (VB) psObj->selected = v through a BASE_OBJECT * (basedef_h's accessors have no Let for it)
Private Sub setObjectSelected(ByVal psObj As Long, ByVal v As Long)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).selected = v
    Case OBJ_STRUCTURE: aStruct(psObj).selected = v
    Case OBJ_FEATURE: aFeature(psObj).selected = v
    End Select
End Sub

Public Sub BaseObjectsController_clearSelection(ByVal c As BaseObjectsController)
    clearSelection
End Sub

Public Sub BaseObjectsController_clearStructureSelection(ByVal c As BaseObjectsController)
    Dim intStrList As OBJ_LIST, it As Long
    If interfaceStructList(intStrList) Then
        it = intStrList.head
        Do While it <> 0
            aStruct(lnObj(it)).selected = 0
            it = lnNext(it)
        Loop
    End If
End Sub

Public Sub BaseObjectsController_selectObject(ByVal c As BaseObjectsController, ByVal object_ As Long, ByVal jumpToHighlightedStatsObject As Boolean)
    If Not (object_ <> 0) Then ASSERT_FAILURE "(object) != nullptr", "src/hci/objects_stats.cpp:48", "selectObject", "%s should not be nullptr", "object": Exit Sub
    setObjectSelected object_, 1       ' object->selected = true;
    triggerEventSelected
    jsDebugSelected object_
    c.setHighlightedObject object_, jumpToHighlightedStatsObject
    c.refresh
End Sub

' virtual void prepareToClose(): the default
Public Sub BaseObjectsController_def_prepareToClose(ByVal c As BaseObjectsController)
    c.clearData
End Sub

Public Sub BaseObjectsController_jumpToObject(ByVal c As BaseObjectsController, ByVal object_ As Long)
    Dim pos As Vector3i
    If Not (object_ <> 0) Then ASSERT_FAILURE "(object) != nullptr", "src/hci/objects_stats.cpp:63", "jumpToObject", "%s should not be nullptr", "object": Exit Sub
    pos = SIMPLE_OBJECT_pos(object_)
    setPlayerPos pos.x, pos.y
    setWarCamActive False
End Sub

Public Sub BaseObjectsController_updateHighlighted(ByVal c As BaseObjectsController)
    Dim findAnySelectedObject As WzFunc, findHighlightedObject As WzFunc, highlighted As Long
    If c.objectsSize() = 0 Then
        c.setHighlightedObject 0, False
        Exit Sub
    End If

    Set findAnySelectedObject = WzFunc_new(FN_findAnySelectedObject, c)

    If c.findObject(findAnySelectedObject) Then
        Exit Sub
    End If

    highlighted = c.getHighlightedObject()
    If highlighted <> 0 Then
        Set findHighlightedObject = WzFunc_new(FN_findHighlightedObject, c, highlighted)

        If c.findObject(findHighlightedObject) Then
            Exit Sub
        End If
    End If

    c.setHighlightedObject c.getObjectAt(0), False
End Sub

Public Sub BaseObjectsController_closeInterface(ByVal c As BaseObjectsController)
    widgScheduleTask WzFunc_new(FN_closeInterface_task)
End Sub

Public Sub BaseObjectsController_closeInterfaceNoAnim(ByVal c As BaseObjectsController)
    widgScheduleTask WzFunc_new(FN_closeInterfaceNoAnim_task)
End Sub

' template<typename A, typename B> static bool findObject(std::vector<A> vector, std::function<bool (B)> iteration):
' the vector is the array items(0 .. count - 1)
Public Function BaseObjectsController_findObject_vector(items() As Long, ByVal count As Long, ByVal iteration As WzFunc) As Boolean
    Dim i As Long
    For i = 0 To count - 1
        If WzFunc_call(iteration, items(i)) Then
            BaseObjectsController_findObject_vector = True
            Exit Function
        End If
    Next

    BaseObjectsController_findObject_vector = False
End Function

' MARK: - BaseStatsController

Public Sub BaseStatsController_displayStatsForm(ByVal c As BaseStatsController)
    Dim newStatsForm As WIDGET
    If widgGetFromID(psWScreen, IDSTAT_FORM) Is Nothing Then
        Set newStatsForm = c.makeStatsForm()
        psWScreen.psForm.attach newStatsForm
        intMode = INT_STAT
        setSecondaryWindowUp True
    End If
End Sub

' To be called from within widget click & run handlers / functions
' (VB) the std::weak_ptr capture is a strong one: the task holds the controller, and the weak
' pointer's lock (the controller still alive: a form still holds it) is "a panel form is still on
' the screen" (the objects and stats forms are what hold a controller besides hci's interfaceController)
Public Sub BaseStatsController_scheduleDisplayStatsForm(ByVal controller As BaseStatsController)
    widgScheduleTask WzFunc_new(FN_scheduleDisplayStatsForm_task, controller)
End Sub

' MARK: - DynamicIntFancyButton

Public Sub DynamicIntFancyButton_construct(ByVal self As WIDGET)
    IntFancyButton_construct self
    self.kind = K_DynamicIntFancyButton
    self.style = self.style Or WFORM_SECONDARY
End Sub

' virtual bool isHighlighted() const = 0
Public Function DynamicIntFancyButton_isHighlighted(ByVal self As WIDGET) As Boolean
    DynamicIntFancyButton_isHighlighted = False      ' (VB) pure virtual: every class made overrides it
End Function

Public Sub DynamicIntFancyButton_updateLayout(ByVal self As WIDGET)
    DynamicIntFancyButton_updateHighlight self
    IntFancyButton_initDisplay self
End Sub

Public Sub DynamicIntFancyButton_clickPrimary(ByVal self As WIDGET)
End Sub

Public Sub DynamicIntFancyButton_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean)
End Sub

Public Sub DynamicIntFancyButton_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim clickAndReleaseOnThisButton As Boolean
    clickAndReleaseOnThisButton = ((W_CLICKFORM_(self).state And WBUT_DOWN) <> 0) ' relies on W_CLICKFORM handling to properly set WBUT_DOWN

    W_CLICKFORM_released self, psContext, key       ' IntFancyButton::released: W_CLICKFORM's

    If Not clickAndReleaseOnThisButton Then
        Exit Sub ' do nothing
    End If

    If key = WKEY_PRIMARY Then
        vt_DynamicIntFancyButton_clickPrimary self
    ElseIf key = WKEY_SECONDARY Then
        vt_DynamicIntFancyButton_clickSecondary self, False
    End If
End Sub

Public Function DynamicIntFancyButton_clickHeld(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY) As Boolean
    If key = WKEY_PRIMARY Then
        vt_DynamicIntFancyButton_clickSecondary self, True
        DynamicIntFancyButton_clickHeld = True
        Exit Function
    End If
    DynamicIntFancyButton_clickHeld = False
End Function

Public Sub DynamicIntFancyButton_updateHighlight(ByVal self As WIDGET)
    If vt_DynamicIntFancyButton_isHighlighted(self) Then
        W_CLICKFORM_(self).state = W_CLICKFORM_(self).state Or WBUT_CLICKLOCK
    Else
        W_CLICKFORM_(self).state = W_CLICKFORM_(self).state And Not WBUT_CLICKLOCK
    End If
End Sub

' MARK: - StatsButton

Public Sub StatsButton_construct(ByVal self As WIDGET)
    DynamicIntFancyButton_construct self
    self.kind = K_StatsButton
End Sub

' virtual BASE_STATS *getStats() = 0
Public Function StatsButton_getStats(ByVal self As WIDGET) As Long
    StatsButton_getStats = 0                        ' (VB) pure virtual
End Function

Public Function StatsButton_getTip(ByVal self As WIDGET) As String
    Dim stats As Long
    stats = vt_StatsButton_getStats(self)
    If stats = 0 Then StatsButton_getTip = "" Else StatsButton_getTip = BASE_STATS_name(stats) ' getLocalizedStatsName(stats)
End Function

Public Sub StatsButton_addProgressBar(ByVal self As WIDGET)
    Dim init As W_BARINIT
    W_BARINIT_construct init
    init.x = STAT_PROGBARX
    init.y = STAT_PROGBARY
    init.width = STAT_PROGBARWIDTH
    init.height = STAT_PROGBARHEIGHT
    init.sCol = psPalette(WZCOL_ACTION_PROGRESS_BAR_MAJOR)
    init.sMinorCol = psPalette(WZCOL_ACTION_PROGRESS_BAR_MINOR)
    init.pTip = "Progress Bar"
    init.style = WBAR_TROUGH Or WIDG_HIDDEN
    init.iRange = GAME_TICKS_PER_SEC
    Set StatsButton_(self).progressBar = W_BARGRAPH_new_init(init)
    self.attach StatsButton_(self).progressBar
    W_BARGRAPH_setBackgroundColour StatsButton_(self).progressBar, psPalette(WZCOL_ACTION_PRODUCTION_RUN_BACKGROUND)
End Sub

' MARK: - ObjectButton

Public Sub ObjectButton_construct(ByVal self As WIDGET)
    Dim zero As Vector2i
    DynamicIntFancyButton_construct self
    self.kind = K_ObjectButton
    ObjectButton_(self).jumpPosition = zero
    IntFancyButton_(self).buttonType = BTMBUTTON
End Sub

Public Function ObjectButton_isHighlighted(ByVal self As WIDGET) As Boolean
    ObjectButton_isHighlighted = False
End Function

' virtual BaseObjectsController &getController() const = 0: the controller object (VB: As Object)
Public Function ObjectButton_getController(ByVal self As WIDGET) As Object
    Set ObjectButton_getController = Nothing         ' (VB) pure virtual
End Function

Public Sub ObjectButton_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean)
    vt_ObjectButton_jump self
End Sub

Public Sub ObjectButton_jump(ByVal self As WIDGET)
    Dim controller As BaseObjectsController, object_ As Long, jumpPosition As Vector2i
    Set controller = vt_ObjectButton_getController(self)
    object_ = controller.getObjectAt(ObjectButton_(self).objectIndex)
    If Not (object_ <> 0) Then ASSERT_FAILURE "(object) != nullptr", "src/hci/objects_stats.cpp:196", "jump", "%s should not be nullptr", "object": Exit Sub

    jumpPosition = ObjectButton_(self).jumpPosition
    If (jumpPosition.x = 0 And jumpPosition.y = 0) Or Not objectOnScreen(object_, 0) Then
        ObjectButton_(self).jumpPosition = getPlayerPos()
        BaseObjectsController_jumpToObject controller, object_
    Else
        setPlayerPos jumpPosition.x, jumpPosition.y
        setWarCamActive False
        jumpPosition.x = 0: jumpPosition.y = 0
        ObjectButton_(self).jumpPosition = jumpPosition
    End If
End Sub

' MARK: - StatsFormButton

Public Sub StatsFormButton_construct(ByVal self As WIDGET)
    StatsButton_construct self
    self.kind = K_StatsFormButton
End Sub

Public Sub StatsFormButton_addCostBar(ByVal self As WIDGET)
    Dim sBarInit As W_BARINIT
    W_BARINIT_construct sBarInit
    sBarInit.x = STAT_TIMEBARX
    sBarInit.y = STAT_TIMEBARY
    sBarInit.width = STAT_PROGBARWIDTH
    sBarInit.height = STAT_PROGBARHEIGHT
    sBarInit.sCol = psPalette(WZCOL_ACTION_PROGRESS_BAR_MAJOR)
    sBarInit.sMinorCol = psPalette(WZCOL_ACTION_PROGRESS_BAR_MINOR)

    sBarInit.iRange = GAME_TICKS_PER_SEC
    Set StatsFormButton_(self).costBar = W_BARGRAPH_new_init(sBarInit)
    self.attach StatsFormButton_(self).costBar
    W_BARGRAPH_setBackgroundColour StatsFormButton_(self).costBar, psPalette(WZCOL_BLACK)
    StatsFormButton_(self).costBar.setTransparentToMouse True
End Sub

Public Function StatsFormButton_getTip(ByVal self As WIDGET) As String
    Dim costString As String, stats As Long, tipString As String
    costString = va("Cost: %u", vt_StatsFormButton_getCost(self)) ' WzString::format(_("Cost: %u"), getCost())
    stats = vt_StatsButton_getStats(self)
    If stats = 0 Then tipString = "" Else tipString = BASE_STATS_name(stats) ' getLocalizedStatsName(stats)
    tipString = tipString & vbLf
    tipString = tipString & costString
    StatsFormButton_getTip = tipString
End Function

' virtual uint32_t getCost() = 0
Public Function StatsFormButton_getCost(ByVal self As WIDGET) As Long
    StatsFormButton_getCost = 0                     ' (VB) pure virtual
End Function

' MARK: - ObjectsForm

Public Sub ObjectsForm_construct(ByVal self As WIDGET)
    IntFormAnimated_construct self, False
    self.kind = K_ObjectsForm
    ObjectsForm_(self).buttonsCount = 0
    ObjectsForm_(self).previousHighlighted = 0
End Sub

' virtual std::shared_ptr<StatsButton> makeStatsButton(size_t buttonIndex) const = 0
Public Function ObjectsForm_makeStatsButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set ObjectsForm_makeStatsButton = Nothing        ' (VB) pure virtual
End Function

' virtual std::shared_ptr<ObjectButton> makeObjectButton(size_t buttonIndex) const = 0
Public Function ObjectsForm_makeObjectButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set ObjectsForm_makeObjectButton = Nothing       ' (VB) pure virtual
End Function

' virtual BaseObjectsController &getController() const = 0
Public Function ObjectsForm_getController(ByVal self As WIDGET) As Object
    Set ObjectsForm_getController = Nothing          ' (VB) pure virtual
End Function

Public Sub ObjectsForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim controller As BaseObjectsController
    If Not IntFormAnimated_isClosing(self) Then
        ObjectsForm_updateButtons self
        Set controller = vt_ObjectsForm_getController(self)
        BaseObjectsController_updateHighlighted controller
        If ObjectsForm_(self).previousHighlighted <> controller.getHighlightedObject() Then
            ObjectsForm_goToHighlightedTab self
        End If
        ObjectsForm_(self).previousHighlighted = controller.getHighlightedObject()
    End If
    IntFormAnimated_display self, xOffset, yOffset
End Sub

Public Sub ObjectsForm_goToHighlightedTab(ByVal self As WIDGET)
    Dim controller As BaseObjectsController, highlighted As Long, i As Long
    Set controller = vt_ObjectsForm_getController(self)
    highlighted = controller.getHighlightedObject()
    For i = 0 To controller.objectsSize() - 1
        If highlighted = controller.getObjectAt(i) Then
            ListTabWidget_goToChildPage ObjectsForm_(self).objectsList, i
            Exit Sub
        End If
    Next
End Sub

Public Sub ObjectsForm_initialize(ByVal self As WIDGET)
    ' creating an obj stat form
    self.id = IDOBJ_FORM
    self.setCalcLayout WzFunc_new(FN_ObjectsForm_calcLayout)

    ObjectsForm_addCloseButton self
    ObjectsForm_addTabList self
End Sub

Public Sub ObjectsForm_addCloseButton(ByVal self As WIDGET)
    Dim init As W_BUTINIT, button As WIDGET
    W_BUTINIT_construct init
    Set init.calcLayout = WzFunc_new(FN_ObjectsForm_closeCalcLayout)
    init.pTip = "Close"
    Set init.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
    init.UserData = PACKDWORD_TRI(0, IMAGE_CLOSEHILIGHT, IMAGE_CLOSE)
    Set button = W_BUTTON_new_init(init)
    self.attach button
    W_BUTTON_addOnClickHandler button, WzFunc_new(FN_ObjectsForm_closeClicked)
End Sub

Public Sub ObjectsForm_addTabList(ByVal self As WIDGET)
    Set ObjectsForm_(self).objectsList = IntListTabWidget_make()
    self.attach ObjectsForm_(self).objectsList
    ObjectsForm_(self).objectsList.id = IDOBJ_TABFORM
    ObjectsForm_(self).objectsList.setCalcLayout WzFunc_new(FN_ObjectsForm_tabListCalcLayout)
End Sub

Public Sub ObjectsForm_updateButtons(ByVal self As WIDGET)
    Dim controller As BaseObjectsController, objectsCount As Long
    Set controller = vt_ObjectsForm_getController(self)
    objectsCount = controller.objectsSize()
    Do While ObjectsForm_(self).buttonsCount < objectsCount
        ObjectsForm_addNewButton self
    Loop

    Do While ObjectsForm_(self).buttonsCount > objectsCount
        ObjectsForm_removeLastButton self
    Loop
End Sub

Public Sub ObjectsForm_addNewButton(ByVal self As WIDGET)
    Dim buttonIndex As Long, buttonHolder As WIDGET, statButton As WIDGET, objectButton As WIDGET
    buttonIndex = ObjectsForm_(self).buttonsCount
    ObjectsForm_(self).buttonsCount = ObjectsForm_(self).buttonsCount + 1
    Set buttonHolder = WIDGET_new()
    ListTabWidget_addWidgetToLayout ObjectsForm_(self).objectsList, buttonHolder

    Set statButton = vt_ObjectsForm_makeStatsButton(self, buttonIndex)
    buttonHolder.attach statButton
    statButton.setGeometry 0, 0, OBJ_BUTWIDTH, OBJ_BUTHEIGHT
    statButton.style = statButton.style Or WFORM_SECONDARY

    Set objectButton = vt_ObjectsForm_makeObjectButton(self, buttonIndex)
    buttonHolder.attach objectButton
    objectButton.setGeometry 0, OBJ_STARTY, OBJ_BUTWIDTH, OBJ_BUTHEIGHT
End Sub

Public Sub ObjectsForm_removeLastButton(ByVal self As WIDGET)
    Dim listWidget As WIDGET, lastButton As WIDGET
    ObjectsForm_(self).buttonsCount = ObjectsForm_(self).buttonsCount - 1
    Set listWidget = ListTabWidget_listWidget(ObjectsForm_(self).objectsList)
    If Not listWidget Is Nothing Then
        Set lastButton = listWidget.childAt(listWidget.childCount() - 1)
        listWidget.detach lastButton
    End If
End Sub

' MARK: - StatsForm

Public Sub StatsForm_construct(ByVal self As WIDGET)
    IntFormAnimated_construct self, False
    self.kind = K_StatsForm
    StatsForm_(self).buttonsCount = 0
End Sub

Public Sub StatsForm_initialize(ByVal self As WIDGET)
    self.id = IDSTAT_FORM
    self.setCalcLayout WzFunc_new(FN_StatsForm_calcLayout)

    StatsForm_addCloseButton self
    StatsForm_addTabList self
End Sub

Public Sub StatsForm_addCloseButton(ByVal self As WIDGET)
    Dim init As W_BUTINIT, button As WIDGET
    W_BUTINIT_construct init
    init.x = STAT_WIDTH - CLOSE_WIDTH
    init.y = 0
    init.width = CLOSE_WIDTH
    init.height = CLOSE_HEIGHT
    init.pTip = "Close"
    Set init.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
    init.UserData = PACKDWORD_TRI(0, IMAGE_CLOSEHILIGHT, IMAGE_CLOSE)
    Set button = W_BUTTON_new_init(init)
    self.attach button
    W_BUTTON_addOnClickHandler button, WzFunc_new(FN_StatsForm_closeClicked)
End Sub

Public Sub StatsForm_addTabList(ByVal self As WIDGET)
    Dim statListWidth As Long
    Set StatsForm_(self).optionList = IntListTabWidget_make()
    self.attach StatsForm_(self).optionList
    StatsForm_(self).optionList.id = IDSTAT_TABFORM
    ListTabWidget_setChildSize StatsForm_(self).optionList, STAT_BUTWIDTH, STAT_BUTHEIGHT
    ListTabWidget_setChildSpacing StatsForm_(self).optionList, STAT_GAP, STAT_GAP
    statListWidth = STAT_BUTWIDTH * 2 + STAT_GAP
    StatsForm_(self).optionList.setGeometry (STAT_WIDTH - statListWidth) \ 2, STAT_TABFORMY, statListWidth, STAT_HEIGHT - STAT_TABFORMY
End Sub

Public Sub StatsForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    vt_StatsForm_updateLayout self
    IntFormAnimated_display self, xOffset, yOffset
End Sub

Public Sub StatsForm_updateLayout(ByVal self As WIDGET)
    StatsForm_updateButtons self
End Sub

' virtual std::shared_ptr<StatsFormButton> makeOptionButton(size_t buttonIndex) const = 0
Public Function StatsForm_makeOptionButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set StatsForm_makeOptionButton = Nothing         ' (VB) pure virtual
End Function

' virtual BaseStatsController &getController() const = 0 (ObjectStatsForm's, = 0 too, returns a
' BaseObjectsStatsController &: VB has no covariant return, the object is the same)
Public Function StatsForm_getController(ByVal self As WIDGET) As Object
    Set StatsForm_getController = Nothing            ' (VB) pure virtual
End Function

' MARK: - BaseObjectsStatsController

Public Sub BaseObjectsStatsController_updateHighlightedObjectStats(ByVal c As BaseObjectsStatsController)
    Dim objects As BaseObjectsController, highlightedObject As Long, size As Long, i As Long
    Set objects = c
    highlightedObject = objects.getHighlightedObject()
    size = objects.objectsSize()

    For i = 0 To size - 1
        If objects.getObjectAt(i) = highlightedObject Then
            c.highlightedObjectStats = objects.getObjectStatsAt(i)
            Exit Sub
        End If
    Next

    c.highlightedObjectStats = 0
End Sub

Public Function BaseObjectsStatsController_isHighlightedObjectStats(ByVal c As BaseObjectsStatsController, ByVal statsIndex As Long) As Boolean
    Dim stats As BaseStatsController
    Set stats = c
    BaseObjectsStatsController_isHighlightedObjectStats = stats.getStatsAt(statsIndex) = c.highlightedObjectStats
End Function

Public Function BaseObjectsStatsController_getHighlightedObjectStats(ByVal c As BaseObjectsStatsController) As Long
    BaseObjectsStatsController_getHighlightedObjectStats = c.highlightedObjectStats
End Function

' MARK: - ObjectStatsForm

Public Sub ObjectStatsForm_construct(ByVal self As WIDGET)
    StatsForm_construct self
    self.kind = K_ObjectStatsForm
    ObjectStatsForm_(self).previousHighlighted = 0
End Sub

Public Sub ObjectStatsForm_updateLayout(ByVal self As WIDGET)
    Dim controller As BaseObjectsStatsController, objects As BaseObjectsController, highlighted As Long
    StatsForm_updateLayout self
    Set controller = vt_StatsForm_getController(self)
    Set objects = controller
    BaseObjectsStatsController_updateHighlightedObjectStats controller
    highlighted = BaseObjectsStatsController_getHighlightedObjectStats(controller)
    If highlighted <> 0 And objects.getQueuedJumpToHighlightedStatsObject() Then
        ObjectStatsForm_goToHighlightedTab self
        objects.clearQueuedJumpToHighlightedStatsObject
    End If
    ObjectStatsForm_(self).previousHighlighted = highlighted
End Sub

Public Sub ObjectStatsForm_goToHighlightedTab(ByVal self As WIDGET)
    Dim controller As BaseObjectsStatsController, stats As BaseStatsController, highlighted As Long, i As Long
    Set controller = vt_StatsForm_getController(self)
    Set stats = controller
    highlighted = BaseObjectsStatsController_getHighlightedObjectStats(controller)
    For i = 0 To stats.statsSize() - 1
        If highlighted = stats.getStatsAt(i) Then
            ListTabWidget_goToChildPage StatsForm_(self).optionList, i
            Exit Sub
        End If
    Next
End Sub

Public Sub StatsForm_updateButtons(ByVal self As WIDGET)
    Dim controller As BaseStatsController, statsCount As Long
    Set controller = vt_StatsForm_getController(self)
    statsCount = controller.statsSize()
    Do While StatsForm_(self).buttonsCount < statsCount
        StatsForm_addNewButton self
    Loop

    Do While StatsForm_(self).buttonsCount > statsCount
        StatsForm_removeLastButton self
    Loop
End Sub

Public Sub StatsForm_addNewButton(ByVal self As WIDGET)
    Dim buttonIndex As Long, button As WIDGET
    buttonIndex = StatsForm_(self).buttonsCount
    StatsForm_(self).buttonsCount = StatsForm_(self).buttonsCount + 1
    Set button = vt_StatsForm_makeOptionButton(self, buttonIndex)
    ListTabWidget_addWidgetToLayout StatsForm_(self).optionList, button
    button.style = button.style Or WFORM_SECONDARY
End Sub

Public Sub StatsForm_removeLastButton(ByVal self As WIDGET)
    Dim listWidget As WIDGET, lastButton As WIDGET
    StatsForm_(self).buttonsCount = StatsForm_(self).buttonsCount - 1
    Set listWidget = ListTabWidget_listWidget(StatsForm_(self).optionList)
    If Not listWidget Is Nothing Then
        Set lastButton = listWidget.childAt(listWidget.childCount() - 1)
        listWidget.detach lastButton
    End If
End Sub

' MARK: - the lambdas

Public Function objects_stats_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, c As BaseObjectsController, object_ As Long, objListWidth As Long
    Select Case f.id
    Case FN_closeInterface_task
        intResetScreen False
    Case FN_closeInterfaceNoAnim_task
        intResetScreen True
    Case FN_findAnySelectedObject
        Set c = f.cap(0)
        object_ = args(0)
        If SIMPLE_OBJECT_died(object_) = 0 And BASE_OBJECT_selected(object_) Then
            c.setHighlightedObject object_, False
            objects_stats_c_fn = True
            Exit Function
        End If

        objects_stats_c_fn = False
    Case FN_findHighlightedObject
        Set c = f.cap(0)
        object_ = args(0)
        If SIMPLE_OBJECT_died(object_) = 0 And object_ = f.cap(1) Then
            c.setHighlightedObject object_, False
            objects_stats_c_fn = True
            Exit Function
        End If

        objects_stats_c_fn = False
    Case FN_scheduleDisplayStatsForm_task
        If Not (widgGetFromID(psWScreen, IDOBJ_FORM) Is Nothing And widgGetFromID(psWScreen, IDSTAT_FORM) Is Nothing) Then ' psWeakController.lock()
            BaseStatsController_displayStatsForm f.cap(0)
        End If
    Case FN_ObjectsForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry OBJ_BACKX, OBJ_BACKY - IIf(getGroupButtonEnabled(), 80, 0), OBJ_BACKWIDTH, OBJ_BACKHEIGHT
    Case FN_ObjectsForm_closeCalcLayout
        Set psWidget = args(0)
        psWidget.setGeometry OBJ_BACKWIDTH - CLOSE_WIDTH, 0, CLOSE_WIDTH, CLOSE_HEIGHT
    Case FN_ObjectsForm_closeClicked
        widgScheduleTask WzFunc_new(FN_ObjectsForm_closeTask)
    Case FN_ObjectsForm_closeTask
        intResetScreen False
        intMode = INT_NORMAL
    Case FN_ObjectsForm_tabListCalcLayout
        Set psWidget = args(0)              ' IntListTabWidget *pObjectsList = static_cast<IntListTabWidget *>(psWidget);
        ' (assert(pObjectsList != nullptr): the C assert, off in a release build)
        ListTabWidget_setChildSize psWidget, OBJ_BUTWIDTH, OBJ_BUTHEIGHT * 2
        ListTabWidget_setChildSpacing psWidget, OBJ_GAP, OBJ_GAP
        objListWidth = OBJ_BUTWIDTH * 5 + STAT_GAP * 4
        psWidget.setGeometry (OBJ_BACKWIDTH - objListWidth) \ 2, OBJ_TABY, objListWidth, OBJ_BACKHEIGHT - OBJ_TABY
    Case FN_StatsForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry STAT_X, STAT_Y, STAT_WIDTH, STAT_HEIGHT
    Case FN_StatsForm_closeClicked
        widgScheduleTask WzFunc_new(FN_StatsForm_closeTask)
    Case FN_StatsForm_closeTask
        intRemoveStats
        intMode = INT_OBJECT
    End Select
End Function
