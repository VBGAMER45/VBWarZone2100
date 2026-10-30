Attribute VB_Name = "build_c"
' src/hci/build.cpp, build.h: the build panel's widgets and BuildController's static members (the
' controller itself is the class BuildController). The make_shared_enabler classes of the make
' functions are the classes themselves; X::make is X_make.
' (VB) widget class BuildObjectButton : ObjectButton
' (VB) widget class BuildStatsButton : StatsButton
' (VB) widget class BuildOptionButton : StatsFormButton
' (VB) widget class BuildObjectsForm : ObjectsForm
' (VB) widget class BuildStatsForm : ObjectStatsForm
Option Explicit

Public BuildController_highlightedBuilder As Long  ' static DROID *BuildController::highlightedBuilder = nullptr
Public BuildController_showFavorites As Boolean    ' static bool BuildController::showFavorites = false

Public Const FN_BuildOptionButton_startBuildPosition As Long = 8100 ' (cap 0: clickedStats, 1: controllerRef)
Public Const FN_BuildOptionButton_toggleFavorites As Long = 8101    ' (cap 0: clickedStats, 1: controllerRef)
Public Const FN_BuildStatsForm_obsoleteClicked As Long = 8102       ' (cap 0: weakController)
Public Const FN_BuildStatsForm_favoriteClicked As Long = 8103       ' (cap 0: weakController)

' static void resetShowFavorites()
Public Sub BuildController_resetShowFavorites()
    BuildController_showFavorites = False
End Sub

' (VB) std::min of two ints
Private Function minL(ByVal a As Long, ByVal b As Long) As Long
    If b < a Then minL = b Else minL = a
End Function

' MARK: - BuildObjectButton

Public Sub BuildObjectButton_construct(ByVal self As WIDGET, ByVal controller As BuildController, ByVal newObjectIndex As Long)
    ObjectButton_construct self
    self.kind = K_BuildObjectButton
    Set BuildObjectButton_(self).controller = controller
    ObjectButton_(self).objectIndex = newObjectIndex
End Sub

Public Function BuildObjectButton_new(ByVal controller As BuildController, ByVal newObjectIndex As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_BuildObjectButton)
    BuildObjectButton_construct widget, controller, newObjectIndex
    Set BuildObjectButton_new = widget
End Function

Public Sub BuildObjectButton_clickPrimary(ByVal self As WIDGET)
    Dim controller As BuildController, droid_ As Long, objectIndex As Long
    Set controller = BuildObjectButton_(self).controller
    objectIndex = ObjectButton_(self).objectIndex
    droid_ = controller.getObjectAt(objectIndex)
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/build.cpp:199", "clickPrimary", "%s should not be nullptr", "droid": Exit Sub

    If keyDown(KEY_LCTRL) Or keyDown(KEY_RCTRL) Or keyDown(KEY_LSHIFT) Or keyDown(KEY_RSHIFT) Then
        controller.toggleBuilderSelection droid_
        Exit Sub
    End If

    BaseObjectsController_clearSelection controller
    BaseObjectsController_selectObject controller, controller.getObjectAt(objectIndex), False
    vt_ObjectButton_jump self

    BaseStatsController_scheduleDisplayStatsForm controller
End Sub

Public Sub BuildObjectButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim droid_ As Long
    vt_DynamicIntFancyButton_updateLayout self
    droid_ = BuildObjectButton_(self).controller.getObjectAt(ObjectButton_(self).objectIndex)
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/build.cpp:219", "display", "%s should not be nullptr", "droid": Exit Sub
    If isDead(droid_) Then
        ASSERT_FAILURE "!isDead(droid)", "src/hci/build.cpp:222", "display", "Droid is dead"
        ' ensure the backing information is refreshed before the next draw
        intRefreshScreen
        Exit Sub
    End If
    IntFancyButton_displayIMD self, AtlasImage_(), ImdObject_Droid(droid_), xOffset, yOffset
    IntFancyButton_displayIfHighlight self, xOffset, yOffset
End Sub

Public Function BuildObjectButton_getController(ByVal self As WIDGET) As Object
    Set BuildObjectButton_getController = BuildObjectButton_(self).controller
End Function

Public Function BuildObjectButton_getTip(ByVal self As WIDGET) As String
    Dim droid_ As Long
    droid_ = BuildObjectButton_(self).controller.getObjectAt(ObjectButton_(self).objectIndex)
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/build.cpp:239", "getTip", "%s should not be nullptr", "droid": BuildObjectButton_getTip = "": Exit Function
    BuildObjectButton_getTip = droidGetName(droid_)
End Function

' MARK: - BuildStatsButton

Public Sub BuildStatsButton_construct(ByVal self As WIDGET)
    StatsButton_construct self
    self.kind = K_BuildStatsButton
End Sub

Public Function BuildStatsButton_make(ByVal controller As BuildController, ByVal objectIndex As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_BuildStatsButton)
    BuildStatsButton_construct widget
    Set BuildStatsButton_(widget).controller = controller
    BuildStatsButton_(widget).objectIndex = objectIndex
    BuildStatsButton_initialize widget
    Set BuildStatsButton_make = widget
End Function

Public Sub BuildStatsButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim stat As Long
    vt_DynamicIntFancyButton_updateLayout self
    stat = vt_StatsButton_getStats(self)
    If stat <> 0 Then
        IntFancyButton_displayIMD self, AtlasImage_(), ImdObject_StructureStat(stat), xOffset, yOffset
    Else
        IntFancyButton_displayIMD self, AtlasImage_(), ImdObject_Component(0), xOffset, yOffset
    End If
    IntFancyButton_displayIfHighlight self, xOffset, yOffset
End Sub

Public Sub BuildStatsButton_updateLayout(ByVal self As WIDGET)
    Dim droid_ As Long
    DynamicIntFancyButton_updateLayout self
    droid_ = BuildStatsButton_(self).controller.getObjectAt(BuildStatsButton_(self).objectIndex)
    BuildStatsButton_updateProgressBar self, droid_
    BuildStatsButton_updateProductionRunSizeLabel self, droid_
End Sub

' STRUCTURE_STATS *getStats() override: as the BASE_STATS * of the virtual
Public Function BuildStatsButton_getStats(ByVal self As WIDGET) As Long
    BuildStatsButton_getStats = structureStatsBase(BuildStatsButton_(self).controller.getObjectStatsAt(BuildStatsButton_(self).objectIndex))
End Function

Public Sub BuildStatsButton_initialize(ByVal self As WIDGET)
    StatsButton_addProgressBar self
    BuildStatsButton_addProductionRunSizeLabel self
End Sub

Public Sub BuildStatsButton_addProductionRunSizeLabel(ByVal self As WIDGET)
    Dim init As W_LABINIT
    W_LABINIT_construct init
    init.style = WIDG_HIDDEN
    init.x = OBJ_TEXTX
    init.y = OBJ_T1TEXTY
    init.width = 16
    init.height = 16
    init.pText = "BUG! (a)"

    Set BuildStatsButton_(self).productionRunSizeLabel = W_LABEL_new_init(init)
    self.attach BuildStatsButton_(self).productionRunSizeLabel
    W_LABEL_setFontColour BuildStatsButton_(self).productionRunSizeLabel, psPalette(WZCOL_ACTION_PRODUCTION_RUN_TEXT)
End Sub

Public Sub BuildStatsButton_updateProgressBar(ByVal self As WIDGET, ByVal droid_ As Long)
    Dim progressBar As WIDGET, structure_ As Long
    Set progressBar = StatsButton_(self).progressBar
    progressBar.hide

    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/build.cpp:313", "updateProgressBar", "%s should not be nullptr", "droid": Exit Sub
    If Not (aDroid(droid_).asBits(COMP_CONSTRUCT) <> 0) Then ASSERT_FAILURE "droid->asBits[COMP_CONSTRUCT]", "src/hci/build.cpp:314", "updateProgressBar", "Invalid droid type"

    structure_ = DroidGetBuildStructure(droid_)
    If structure_ <> 0 Then
        'show progress of build
        If aStruct(structure_).currentBuildPts <> 0 Then
            formatTime progressBar, aStruct(structure_).currentBuildPts, structureBuildPointsToCompletion(structure_), aStruct(structure_).lastBuildRate, "Build Progress"
        Else
            formatPower progressBar, checkPowerRequest(structure_), asStructureStats(aStruct(structure_).pStructureType).powerToBuild
        End If
    End If
End Sub

' updateProductionRunSizeLabel's lambda processOrder ([&]: stats, count)
Private Function BuildStatsButton_processOrder(order As DROID_ORDER_DATA, stats As Long, count As Long) As Boolean
    Dim newStats As Long, deltaCount As Long, target As Long
    newStats = -1
    deltaCount = 0
    Select Case order.typ
    Case DORDER_BUILD, DORDER_LINEBUILD
        newStats = order.psStats
        If order.typ = DORDER_LINEBUILD Then deltaCount = 1 + (Abs(order.pos.x - order.pos2.x) + Abs(order.pos.y - order.pos2.y)) \ TILE_UNITS Else deltaCount = 1
    Case DORDER_HELPBUILD
        target = castStructure(order.psObj)
        If target <> 0 Then
            newStats = aStruct(target).pStructureType
            deltaCount = 1
        End If
    Case Else
        BuildStatsButton_processOrder = False
        Exit Function
    End Select
    If newStats <> -1 And (stats = -1 Or stats = newStats) Then
        stats = newStats
        count = count + deltaCount
        BuildStatsButton_processOrder = True
        Exit Function
    End If
    BuildStatsButton_processOrder = False
End Function

Public Sub BuildStatsButton_updateProductionRunSizeLabel(ByVal self As WIDGET, ByVal droid_ As Long)
    Dim remaining As Long, stats As Long, count As Long, i As Long, label As WIDGET
    remaining = -1

    stats = -1                          ' STRUCTURE_STATS const *stats = nullptr
    count = 0

    If droid_ <> 0 Then
        If BuildStatsButton_processOrder(aDroid(droid_).order, stats, count) Then
            For i = 0 To aDroid(droid_).asOrderList_size - 1
                If Not BuildStatsButton_processOrder(aDroid(droid_).asOrderList(i), stats, count) Then
                    Exit For
                End If
            Next
        End If
    End If
    If count > 1 Then
        remaining = count
    End If

    Set label = BuildStatsButton_(self).productionRunSizeLabel
    If remaining <> -1 Then
        label.setString va("%d", remaining)
        label.show
    Else
        label.hide
    End If
End Sub

Public Function BuildStatsButton_isHighlighted(ByVal self As WIDGET) As Boolean
    Dim controller As BuildController, droid_ As Long
    Set controller = BuildStatsButton_(self).controller
    droid_ = controller.getObjectAt(BuildStatsButton_(self).objectIndex)
    If droid_ <> 0 Then BuildStatsButton_isHighlighted = (aDroid(droid_).selected <> 0 Or droid_ = controller.getHighlightedObject())
End Function

Public Sub BuildStatsButton_clickPrimary(ByVal self As WIDGET)
    Dim controller As BuildController, droid_ As Long
    Set controller = BuildStatsButton_(self).controller
    droid_ = controller.getObjectAt(BuildStatsButton_(self).objectIndex)
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/build.cpp:400", "clickPrimary", "%s should not be nullptr", "droid": Exit Sub

    If keyDown(KEY_LCTRL) Or keyDown(KEY_RCTRL) Or keyDown(KEY_LSHIFT) Or keyDown(KEY_RSHIFT) Then
        controller.toggleBuilderSelection droid_
    Else
        BaseObjectsController_clearSelection controller
        BaseObjectsController_selectObject controller, droid_, True
    End If

    BaseStatsController_scheduleDisplayStatsForm controller
End Sub

Public Sub BuildStatsButton_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean)
    Dim controller As BuildController, droid_ As Long, highlighted As Long
    Set controller = BuildStatsButton_(self).controller
    droid_ = controller.getObjectAt(BuildStatsButton_(self).objectIndex)
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/build.cpp:418", "clickSecondary", "%s should not be nullptr", "droid": Exit Sub
    highlighted = controller.getHighlightedObject()

    ' prevent highlighting a builder when another builder is already selected
    If droid_ = highlighted Then
        controller.setHighlightedObject droid_, True
        BaseStatsController_scheduleDisplayStatsForm controller
    ElseIf aDroid(highlighted).selected = 0 Then
        controller.setHighlightedObject droid_, True
        BaseStatsController_scheduleDisplayStatsForm controller
    End If
End Sub

' MARK: - BuildOptionButton

Public Sub BuildOptionButton_construct(ByVal self As WIDGET)
    StatsFormButton_construct self
    self.kind = K_BuildOptionButton
End Sub

Public Function BuildOptionButton_make(ByVal controller As BuildController, ByVal buildOptionIndex As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_BuildOptionButton)
    BuildOptionButton_construct widget
    Set BuildOptionButton_(widget).controller = controller
    BuildOptionButton_(widget).buildOptionIndex = buildOptionIndex
    BuildOptionButton_initialize widget
    Set BuildOptionButton_make = widget
End Function

Public Sub BuildOptionButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim stat As Long
    vt_DynamicIntFancyButton_updateLayout self
    stat = vt_StatsButton_getStats(self)
    If Not (stat <> 0) Then ASSERT_FAILURE "(stat) != nullptr", "src/hci/build.cpp:456", "display", "%s should not be nullptr", "stat": Exit Sub

    IntFancyButton_displayIMD self, AtlasImage_(), ImdObject_StructureStat(stat), xOffset, yOffset
    IntFancyButton_displayIfHighlight self, xOffset, yOffset
    If asStructureStats(castStructureStats(stat)).isFavorite And Not BuildOptionButton_(self).controller.shouldShowFavorites() Then
        iV_DrawImage IntImages, IMAGE_ALLY_RESEARCH_TC, xOffset + self.x() + 44, yOffset + self.y() + 3
    End If
End Sub

' STRUCTURE_STATS *getStats() override: as the BASE_STATS * of the virtual
Public Function BuildOptionButton_getStats(ByVal self As WIDGET) As Long
    BuildOptionButton_getStats = structureStatsBase(BuildOptionButton_(self).controller.getStatsAt(BuildOptionButton_(self).buildOptionIndex))
End Function

Public Sub BuildOptionButton_initialize(ByVal self As WIDGET)
    StatsFormButton_addCostBar self
End Sub

Public Function BuildOptionButton_isHighlighted(ByVal self As WIDGET) As Boolean
    BuildOptionButton_isHighlighted = BaseObjectsStatsController_isHighlightedObjectStats(BuildOptionButton_(self).controller, BuildOptionButton_(self).buildOptionIndex)
End Function

Public Sub BuildOptionButton_updateLayout(ByVal self As WIDGET)
    DynamicIntFancyButton_updateLayout self

    If self.isMouseOverWidget() Then
        intSetShadowPower vt_StatsFormButton_getCost(self)
    End If

    W_BARGRAPH_(StatsFormButton_(self).costBar).majorSize = minL(100, vt_StatsFormButton_getCost(self) \ POWERPOINTS_DROIDDIV)
End Sub

Public Function BuildOptionButton_getCost(ByVal self As WIDGET) As Long
    Dim psStats As Long
    psStats = castStructureStats(vt_StatsButton_getStats(self))
    If psStats <> -1 Then BuildOptionButton_getCost = asStructureStats(psStats).powerToBuild Else BuildOptionButton_getCost = 0
End Function

Public Sub BuildOptionButton_clickPrimary(ByVal self As WIDGET)
    Dim controller As BuildController, clickedStats As Long, controllerRef As BuildController
    Set controller = BuildOptionButton_(self).controller
    clickedStats = controller.getStatsAt(BuildOptionButton_(self).buildOptionIndex)
    If Not (clickedStats <> -1) Then ASSERT_FAILURE "(clickedStats) != nullptr", "src/hci/build.cpp:503", "clickPrimary", "%s should not be nullptr", "clickedStats": Exit Sub

    Set controllerRef = controller
    widgScheduleTask WzFunc_new(FN_BuildOptionButton_startBuildPosition, clickedStats, controllerRef)
End Sub

Public Sub BuildOptionButton_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean)
    Dim controller As BuildController, clickedStats As Long, controllerRef As BuildController
    Set controller = BuildOptionButton_(self).controller
    clickedStats = controller.getStatsAt(BuildOptionButton_(self).buildOptionIndex)
    If Not (clickedStats <> -1) Then ASSERT_FAILURE "(clickedStats) != nullptr", "src/hci/build.cpp:514", "clickSecondary", "%s should not be nullptr", "clickedStats": Exit Sub

    Set controllerRef = controller
    widgScheduleTask WzFunc_new(FN_BuildOptionButton_toggleFavorites, clickedStats, controllerRef)
End Sub

' MARK: - BuildObjectsForm

Public Sub BuildObjectsForm_construct(ByVal self As WIDGET)
    ObjectsForm_construct self
    self.kind = K_BuildObjectsForm
End Sub

Public Function BuildObjectsForm_make(ByVal controller As BuildController) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_BuildObjectsForm)
    BuildObjectsForm_construct widget
    Set BuildObjectsForm_(widget).controller = controller
    ObjectsForm_initialize widget
    Set BuildObjectsForm_make = widget
End Function

Public Function BuildObjectsForm_makeStatsButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set BuildObjectsForm_makeStatsButton = BuildStatsButton_make(BuildObjectsForm_(self).controller, buttonIndex)
End Function

Public Function BuildObjectsForm_makeObjectButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set BuildObjectsForm_makeObjectButton = BuildObjectButton_new(BuildObjectsForm_(self).controller, buttonIndex)
End Function

Public Function BuildObjectsForm_getController(ByVal self As WIDGET) As Object
    Set BuildObjectsForm_getController = BuildObjectsForm_(self).controller
End Function

' MARK: - BuildStatsForm

Public Sub BuildStatsForm_construct(ByVal self As WIDGET)
    ObjectStatsForm_construct self
    self.kind = K_BuildStatsForm
End Sub

Public Function BuildStatsForm_make(ByVal controller As BuildController) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_BuildStatsForm)
    BuildStatsForm_construct widget
    Set BuildStatsForm_(widget).controller = controller
    vt_StatsForm_initialize widget
    Set BuildStatsForm_make = widget
End Function

Public Function BuildStatsForm_makeOptionButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set BuildStatsForm_makeOptionButton = BuildOptionButton_make(BuildStatsForm_(self).controller, buttonIndex)
End Function

Public Function BuildStatsForm_getController(ByVal self As WIDGET) As Object
    Set BuildStatsForm_getController = BuildStatsForm_(self).controller
End Function

Public Sub BuildStatsForm_initialize(ByVal self As WIDGET)
    StatsForm_initialize self                       ' BaseWidget::initialize(): ObjectStatsForm's is StatsForm's
    BuildStatsForm_addObsoleteButton self
    BuildStatsForm_addFavoriteButton self
End Sub

Public Sub BuildStatsForm_addObsoleteButton(ByVal self As WIDGET)
    Dim obsoleteButton As WIDGET, weakController As BuildController
    Set obsoleteButton = MultipleChoiceButton_new()
    Set BuildStatsForm_(self).obsoleteButton = obsoleteButton
    self.attach obsoleteButton
    obsoleteButton.style = obsoleteButton.style Or WBUT_SECONDARY
    MultipleChoiceButton_setChoice obsoleteButton, IIf(BuildStatsForm_(self).controller.shouldShowRedundantDesign(), 1, 0)
    MultipleChoiceButton_setImages obsoleteButton, 0, W_BUTTON_Images_(AtlasImage_(IntImages, IMAGE_OBSOLETE_HIDE_UP), AtlasImage_(IntImages, IMAGE_OBSOLETE_HIDE_UP), AtlasImage_(IntImages, IMAGE_OBSOLETE_HIDE_HI), AtlasImage_())
    MultipleChoiceButton_setTip_choice obsoleteButton, 0, "Hiding Obsolete Tech"
    MultipleChoiceButton_setImages obsoleteButton, 1, W_BUTTON_Images_(AtlasImage_(IntImages, IMAGE_OBSOLETE_SHOW_UP), AtlasImage_(IntImages, IMAGE_OBSOLETE_SHOW_UP), AtlasImage_(IntImages, IMAGE_OBSOLETE_SHOW_HI), AtlasImage_())
    MultipleChoiceButton_setTip_choice obsoleteButton, 1, "Showing Obsolete Tech"
    obsoleteButton.move 4 + AtlasImage_width(IntImages, IMAGE_FDP_UP) + 4, STAT_SLDY

    Set weakController = BuildStatsForm_(self).controller ' std::weak_ptr<BuildController>(controller): a strong reference
    W_BUTTON_addOnClickHandler obsoleteButton, WzFunc_new(FN_BuildStatsForm_obsoleteClicked, weakController)
End Sub

Public Sub BuildStatsForm_addFavoriteButton(ByVal self As WIDGET)
    Dim favoriteButton As WIDGET, weakController As BuildController
    Set favoriteButton = MultipleChoiceButton_new()
    Set BuildStatsForm_(self).favoriteButton = favoriteButton
    self.attach favoriteButton
    favoriteButton.style = favoriteButton.style Or WBUT_SECONDARY
    MultipleChoiceButton_setChoice favoriteButton, IIf(BuildStatsForm_(self).controller.shouldShowFavorites(), 1, 0)
    MultipleChoiceButton_setImages favoriteButton, 0, W_BUTTON_Images_(AtlasImage_(IntImages, IMAGE_ALLY_RESEARCH), AtlasImage_(IntImages, IMAGE_ALLY_RESEARCH), AtlasImage_(IntImages, IMAGE_ALLY_RESEARCH), AtlasImage_())
    MultipleChoiceButton_setTip_choice favoriteButton, 0, "Showing All Tech" & vbLf & "Right-click to add to Favorites"
    MultipleChoiceButton_setImages favoriteButton, 1, W_BUTTON_Images_(AtlasImage_(IntImages, IMAGE_ALLY_RESEARCH_TC), AtlasImage_(IntImages, IMAGE_ALLY_RESEARCH_TC), AtlasImage_(IntImages, IMAGE_ALLY_RESEARCH_TC), AtlasImage_())
    MultipleChoiceButton_setTip_choice favoriteButton, 1, "Showing Only Favorite Tech" & vbLf & "Right-click to remove from Favorites"
    favoriteButton.move 4 * 2 + AtlasImage_width(IntImages, IMAGE_FDP_UP) * 2 + 4 * 2, STAT_SLDY

    Set weakController = BuildStatsForm_(self).controller ' std::weak_ptr<BuildController>(controller): a strong reference
    W_BUTTON_addOnClickHandler favoriteButton, WzFunc_new(FN_BuildStatsForm_favoriteClicked, weakController)
End Sub

' MARK: - the lambdas

Public Function build_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim controllerRef As BuildController, buildController As BuildController, button As WIDGET, newValue As Boolean
    Select Case f.id
    Case FN_BuildOptionButton_startBuildPosition
        Set controllerRef = f.cap(1)
        controllerRef.startBuildPosition f.cap(0)
    Case FN_BuildOptionButton_toggleFavorites
        Set controllerRef = f.cap(1)
        controllerRef.toggleFavorites f.cap(0)
    Case FN_BuildStatsForm_obsoleteClicked
        Set buildController = f.cap(0)  ' weakController.lock()
        If Not buildController Is Nothing Then
            Set button = args(0)        ' auto &_obsoleteButton = static_cast<MultipleChoiceButton &>(button);
            newValue = Not (MultipleChoiceButton_getChoice(button) <> 0)
            buildController.setShouldShowRedundantDesign newValue
            MultipleChoiceButton_setChoice button, IIf(newValue, 1, 0)
        End If
    Case FN_BuildStatsForm_favoriteClicked
        Set buildController = f.cap(0)  ' weakController.lock()
        If Not buildController Is Nothing Then
            Set button = args(0)        ' auto &_favoriteButton = static_cast<MultipleChoiceButton &>(button);
            newValue = Not (MultipleChoiceButton_getChoice(button) <> 0)
            buildController.setShouldShowFavorite newValue
            MultipleChoiceButton_setChoice button, IIf(newValue, 1, 0)
        End If
    End Select
End Function
