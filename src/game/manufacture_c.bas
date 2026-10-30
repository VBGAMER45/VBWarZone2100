Attribute VB_Name = "manufacture_c"
' src/hci/manufacture.cpp, manufacture.h: the manufacture panel's widgets, the factory helpers and
' ManufactureController's static member (the controller itself is the class
' ManufactureController). ManufactureStatsForm's nested classes DeliveryPointButton,
' LoopProductionButton and LoopProductionLabel are widget classes of their own names. The static
' functions the controller class calls (compareFactories) are Public.
' (VB) widget class ManufactureObjectButton : ObjectButton
' (VB) widget class ManufactureStatsButton : StatsButton
' (VB) widget class ManufactureOptionButton : StatsFormButton
' (VB) widget class ManufactureObjectsForm : ObjectsForm
' (VB) widget class ManufactureStatsForm : ObjectStatsForm
' (VB) widget class DeliveryPointButton : W_BUTTON
' (VB) widget class LoopProductionButton : W_BUTTON
' (VB) widget class LoopProductionLabel : W_LABEL
Option Explicit

Public ManufactureController_highlightedFactory As Long ' static STRUCTURE *ManufactureController::highlightedFactory = nullptr
Private Const nullBrainComponent As Long = 0    ' "0" to reference "NULLxxx" components
Private Const commandBrainComponent As Long = 1 ' hmm there is only 1 "CommandBrain01" anyway..

Public Const FN_ManufactureOptionButton_adjust As Long = 8200       ' (cap 0: add, 1: clickedStats, 2: controllerRef)
Public Const FN_ManufactureStatsForm_obsoleteClicked As Long = 8201 ' (cap 0: weakController)
Public Const FN_DeliveryPointButton_clicked As Long = 8202          ' (cap 0: weakController)

' (VB) std::min of two ints
Private Function minL(ByVal a As Long, ByVal b As Long) As Long
    If b < a Then minL = b Else minL = a
End Function

' FACTORY *: the structure's handle (its pFunctionality), 0 = nullptr
Public Function getFactoryOrNullptr(ByVal factory_ As Long) As Long
    Dim ok As Boolean
    If factory_ <> 0 Then ok = STRUCTURE_isFactory(factory_)
    If Not ok Then ASSERT_FAILURE "factory && factory->isFactory()", "src/hci/manufacture.cpp:36", "getFactoryOrNullptr", "Invalid factory pointer": getFactoryOrNullptr = 0: Exit Function
    getFactoryOrNullptr = aStruct(factory_).pFunctionality
End Function

Private Function getProductionLoops(ByVal structure_ As Long) As Long
    Dim psFactory As Long
    If structure_ = 0 Then
        getProductionLoops = 0
        Exit Function
    End If

    psFactory = getFactoryOrNullptr(structure_)
    If Not (psFactory <> 0) Then ASSERT_FAILURE "(psFactory) != nullptr", "src/hci/manufacture.cpp:48", "getProductionLoops", "%s should not be nullptr", "psFactory": getProductionLoops = 0: Exit Function
    getProductionLoops = aStruct(psFactory).functionality.factory.productionLoops
End Function

Private Function makeProductionRunSizeLabel() As WIDGET
    Dim init As W_LABINIT, label As WIDGET
    W_LABINIT_construct init
    init.x = OBJ_TEXTX
    init.y = OBJ_T1TEXTY
    init.width = 16
    init.height = 16
    Set label = W_LABEL_new_init(init)
    label.setTransparentToMouse True
    Set makeProductionRunSizeLabel = label
End Function

' static inline bool compareFactories(STRUCTURE *a, STRUCTURE *b) (Public: ManufactureController's sort)
Public Function compareFactories(ByVal a As Long, ByVal b As Long) As Boolean
    Dim x As Long, y As Long
    If a = 0 Or b = 0 Then
        compareFactories = (a <> 0 And b = 0)   ' (a == nullptr) < (b == nullptr): false < true
        Exit Function
    End If
    x = getFactoryOrNullptr(a)
    If Not (x <> 0) Then ASSERT_FAILURE "(x) != nullptr", "src/hci/manufacture.cpp:124", "compareFactories", "%s should not be nullptr", "x": compareFactories = False: Exit Function
    y = getFactoryOrNullptr(b)
    If Not (y <> 0) Then ASSERT_FAILURE "(y) != nullptr", "src/hci/manufacture.cpp:126", "compareFactories", "%s should not be nullptr", "y": compareFactories = False: Exit Function
    If Not (aStruct(x).functionality.factory.psAssemblyPoint <> 0) Then ASSERT_FAILURE "(x->psAssemblyPoint) != nullptr", "src/hci/manufacture.cpp:127", "compareFactories", "%s should not be nullptr", "x->psAssemblyPoint": compareFactories = False: Exit Function
    If Not (aStruct(y).functionality.factory.psAssemblyPoint <> 0) Then ASSERT_FAILURE "(y->psAssemblyPoint) != nullptr", "src/hci/manufacture.cpp:128", "compareFactories", "%s should not be nullptr", "y->psAssemblyPoint": compareFactories = False: Exit Function
    If aFlagPos(aStruct(x).functionality.factory.psAssemblyPoint).factoryType <> aFlagPos(aStruct(y).functionality.factory.psAssemblyPoint).factoryType Then
        compareFactories = aFlagPos(aStruct(x).functionality.factory.psAssemblyPoint).factoryType < aFlagPos(aStruct(y).functionality.factory.psAssemblyPoint).factoryType
        Exit Function
    End If

    compareFactories = aFlagPos(aStruct(x).functionality.factory.psAssemblyPoint).factoryInc < aFlagPos(aStruct(y).functionality.factory.psAssemblyPoint).factoryInc
End Function

' MARK: - ManufactureObjectButton

Public Sub ManufactureObjectButton_construct(ByVal self As WIDGET)
    ObjectButton_construct self
    self.kind = K_ManufactureObjectButton
End Sub

Public Function ManufactureObjectButton_make(ByVal controller As ManufactureController, ByVal objectIndex As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_ManufactureObjectButton)
    ManufactureObjectButton_construct widget
    Set ManufactureObjectButton_(widget).controller = controller
    ObjectButton_(widget).objectIndex = objectIndex
    ManufactureObjectButton_initialize widget
    Set ManufactureObjectButton_make = widget
End Function

Public Sub ManufactureObjectButton_jump(ByVal self As WIDGET)
    If Not offWorldKeepLists Then
        ObjectButton_jump self
    End If
End Sub

Public Sub ManufactureObjectButton_clickPrimary(ByVal self As WIDGET)
    Dim controller As ManufactureController
    Set controller = ManufactureObjectButton_(self).controller
    BaseObjectsController_clearStructureSelection controller
    BaseObjectsController_selectObject controller, controller.getObjectAt(ObjectButton_(self).objectIndex), False
    vt_ObjectButton_jump self
    BaseStatsController_scheduleDisplayStatsForm controller
End Sub

Public Sub ManufactureObjectButton_initialize(ByVal self As WIDGET)
    With ManufactureObjectButton_(self)
        Set .factoryNumberLabel = W_LABEL_new()
        self.attach .factoryNumberLabel
        Set .factoryAssignGroupLabel = W_LABEL_new()
        self.attach .factoryAssignGroupLabel
        .factoryNumberLabel.setGeometry OBJ_TEXTX, OBJ_B1TEXTY, 16, 16
        .factoryAssignGroupLabel.setGeometry OBJ_TEXTX + 48, OBJ_B1TEXTY, 16, 16
        W_LABEL_setFontColour .factoryAssignGroupLabel, pal_RGBA(255, 220, 115, 255) ' gold
    End With
End Sub

Public Sub ManufactureObjectButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim factory_ As Long
    vt_DynamicIntFancyButton_updateLayout self
    factory_ = ManufactureObjectButton_(self).controller.getObjectAt(ObjectButton_(self).objectIndex)
    If Not (factory_ <> 0) Then ASSERT_FAILURE "(factory) != nullptr", "src/hci/manufacture.cpp:254", "display", "%s should not be nullptr", "factory": Exit Sub
    If isDead(factory_) Then
        ASSERT_FAILURE "!isDead(factory)", "src/hci/manufacture.cpp:257", "display", "Factory is dead"
        ' ensure the backing information is refreshed before the next draw
        intRefreshScreen
        Exit Sub
    End If
    IntFancyButton_displayIMD self, AtlasImage_(), ImdObject_Structure(factory_), xOffset, yOffset
    IntFancyButton_displayIfHighlight self, xOffset, yOffset
End Sub

Public Sub ManufactureObjectButton_updateLayout(ByVal self As WIDGET)
    Dim psStruct As Long, factory_ As Long, psAssemblyPoint As Long
    DynamicIntFancyButton_updateLayout self
    psStruct = ManufactureObjectButton_(self).controller.getObjectAt(ObjectButton_(self).objectIndex)
    factory_ = getFactoryOrNullptr(psStruct)
    If Not (factory_ <> 0) Then ASSERT_FAILURE "(factory) != nullptr", "src/hci/manufacture.cpp:271", "updateLayout", "%s should not be nullptr", "factory": Exit Sub
    psAssemblyPoint = aStruct(factory_).functionality.factory.psAssemblyPoint
    If psAssemblyPoint = 0 Then
        ManufactureObjectButton_(self).factoryNumberLabel.setString ""
    Else
        ManufactureObjectButton_(self).factoryNumberLabel.setString va("%u", aFlagPos(psAssemblyPoint).factoryInc + 1)
    End If
    If aStruct(psStruct).productToGroup <> UBYTE_MAX Then
        ManufactureObjectButton_(self).factoryAssignGroupLabel.setString va("%u", aStruct(psStruct).productToGroup)
    Else
        ManufactureObjectButton_(self).factoryAssignGroupLabel.setString ""
    End If
End Sub

Public Function ManufactureObjectButton_getTip(ByVal self As WIDGET) As String
    Dim factory_ As Long
    factory_ = ManufactureObjectButton_(self).controller.getObjectAt(ObjectButton_(self).objectIndex)
    If Not (factory_ <> 0) Then ASSERT_FAILURE "(factory) != nullptr", "src/hci/manufacture.cpp:293", "getTip", "%s should not be nullptr", "factory": ManufactureObjectButton_getTip = "": Exit Function
    ManufactureObjectButton_getTip = BASE_STATS_name(structureStatsBase(aStruct(factory_).pStructureType)) ' getLocalizedStatsName(factory->pStructureType)
End Function

Public Function ManufactureObjectButton_getController(ByVal self As WIDGET) As Object
    Set ManufactureObjectButton_getController = ManufactureObjectButton_(self).controller
End Function

' MARK: - ManufactureStatsButton

Public Sub ManufactureStatsButton_construct(ByVal self As WIDGET)
    StatsButton_construct self
    self.kind = K_ManufactureStatsButton
End Sub

Public Function ManufactureStatsButton_make(ByVal controller As ManufactureController, ByVal objectIndex As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_ManufactureStatsButton)
    ManufactureStatsButton_construct widget
    Set ManufactureStatsButton_(widget).controller = controller
    ManufactureStatsButton_(widget).objectIndex = objectIndex
    ManufactureStatsButton_initialize widget
    Set ManufactureStatsButton_make = widget
End Function

Public Sub ManufactureStatsButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim factory_ As Long, production As Long, productionPending As Boolean, objectImage As ImdObject
    vt_DynamicIntFancyButton_updateLayout self

    factory_ = ManufactureStatsButton_(self).controller.getObjectAt(ManufactureStatsButton_(self).objectIndex)
    production = vt_StatsButton_getStats(self)
    If factory_ <> 0 Then productionPending = StructureIsManufacturingPending(factory_)
    If productionPending And production <> 0 Then objectImage = ImdObject_DroidTemplate(production) Else objectImage = ImdObject_Component(0)
    IntFancyButton_displayIMD self, AtlasImage_(), objectImage, xOffset, yOffset

    If productionPending And StructureIsOnHoldPending(factory_) Then
        If ((realTime \ 250) Mod 2) = 0 Then
            iV_DrawImage IntImages, IMAGE_BUT0_DOWN, xOffset + self.x(), yOffset + self.y()
        Else
            iV_DrawImage IntImages, IMAGE_BUT_HILITE, xOffset + self.x(), yOffset + self.y()
        End If
    Else
        IntFancyButton_displayIfHighlight self, xOffset, yOffset
    End If
    ManufactureStatsButton_drawNextDroidRank self, xOffset, yOffset
End Sub

' show a little icon on the bottom, indicating what rank next unit will be
' remember that when it's a commander, same rank will require twice experience
Public Sub ManufactureStatsButton_drawNextDroidRank(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim factory_ As Long, player As Long, exp_ As Long, lvl As Long, psSubject As Long, expgfx As Long
    factory_ = ManufactureStatsButton_(self).controller.getObjectAt(ManufactureStatsButton_(self).objectIndex)
    If factory_ = 0 Then
        Exit Sub
    End If
    player = aStruct(factory_).player
    exp_ = getTopExperience(player)
    lvl = 0
    psSubject = aStruct(aStruct(factory_).pFunctionality).functionality.factory.psSubject
    If psSubject = 0 Then
        ' not showing icon when nothing is being built
        Exit Sub
    End If
    If apsTemplates(psSubject).droidType = DROID_COMMAND Or apsTemplates(psSubject).droidType = DROID_SENSOR Then
        lvl = getDroidLevel_exp(exp_, player, commandBrainComponent)
    ElseIf apsTemplates(psSubject).droidType = DROID_CONSTRUCT _
    Or apsTemplates(psSubject).droidType = DROID_CYBORG_CONSTRUCT _
    Or apsTemplates(psSubject).droidType = DROID_CYBORG_REPAIR _
    Or apsTemplates(psSubject).droidType = DROID_REPAIR _
    Or apsTemplates(psSubject).droidType = DROID_TRANSPORTER _
    Or apsTemplates(psSubject).droidType = DROID_SUPERTRANSPORTER Then
        lvl = 0
    Else
        lvl = getDroidLevel_exp(exp_, player, nullBrainComponent)
    End If
    expgfx = getDroidRankGraphicFromLevel(lvl)
    If expgfx <> UDWORD_MAX Then
        ' FIXME: use offsets relative to template positon, not hardcoded values ?
        iV_DrawImage IntImages, expgfx And &HFFFF&, xOffset + 45, yOffset + 4
    End If
End Sub

Public Sub ManufactureStatsButton_updateLayout(ByVal self As WIDGET)
    Dim factory_ As Long
    DynamicIntFancyButton_updateLayout self
    factory_ = ManufactureStatsButton_(self).controller.getObjectAt(ManufactureStatsButton_(self).objectIndex)
    ManufactureStatsButton_updateProgressBar self, factory_
    ManufactureStatsButton_updateProductionRunSizeLabel self, factory_, FactoryGetTemplate(StructureGetFactory(factory_))
End Sub

' DROID_TEMPLATE *getStats() override: as the BASE_STATS * of the virtual
Public Function ManufactureStatsButton_getStats(ByVal self As WIDGET) As Long
    Dim psTemplate As Long
    psTemplate = ManufactureStatsButton_(self).controller.getObjectStatsAt(ManufactureStatsButton_(self).objectIndex)
    If psTemplate <> 0 Then ManufactureStatsButton_getStats = STAT_TEMPLATE + psTemplate
End Function

Public Sub ManufactureStatsButton_initialize(ByVal self As WIDGET)
    StatsButton_addProgressBar self
    ManufactureStatsButton_addProductionRunSizeLabel self
End Sub

Public Sub ManufactureStatsButton_addProductionRunSizeLabel(ByVal self As WIDGET)
    Set ManufactureStatsButton_(self).productionRunSizeLabel = makeProductionRunSizeLabel()
    self.attach ManufactureStatsButton_(self).productionRunSizeLabel
    W_LABEL_setFontColour ManufactureStatsButton_(self).productionRunSizeLabel, psPalette(WZCOL_ACTION_PRODUCTION_RUN_TEXT)
End Sub

' droidTemplate: a DROID_TEMPLATE * (0 = nullptr)
Public Sub ManufactureStatsButton_updateProductionRunSizeLabel(ByVal self As WIDGET, ByVal factory_ As Long, ByVal droidTemplate As Long)
    Dim production As ProductionRunEntry, productionRemaining As Long, pending As Boolean
    production = getProduction(factory_, droidTemplate)
    productionRemaining = ProductionRunEntry_numRemaining(production)
    If productionRemaining > 0 And factory_ <> 0 Then pending = StructureIsManufacturingPending(factory_)
    If pending Then
        ManufactureStatsButton_(self).productionRunSizeLabel.setString va("%d", productionRemaining)
        ManufactureStatsButton_(self).productionRunSizeLabel.show
    Else
        ManufactureStatsButton_(self).productionRunSizeLabel.hide
    End If
End Sub

Public Sub ManufactureStatsButton_updateProgressBar(ByVal self As WIDGET, ByVal factory_ As Long)
    Dim progressBar As WIDGET, manufacture As Long, buildPointsTotal As Long, buildRate As Long, neededPower As Long, powerToBuild As Long
    Set progressBar = StatsButton_(self).progressBar
    progressBar.hide

    If Not (factory_ <> 0) Then ASSERT_FAILURE "(factory) != nullptr", "src/hci/manufacture.cpp:435", "updateProgressBar", "%s should not be nullptr", "factory": Exit Sub
    If Not (Not isDead(factory_)) Then ASSERT_FAILURE "!isDead(factory)", "src/hci/manufacture.cpp:436", "updateProgressBar", "Factory is dead": Exit Sub

    If Not StructureIsManufacturingPending(factory_) Then
        Exit Sub
    End If

    manufacture = StructureGetFactory(factory_)
    If Not (manufacture <> 0) Then ASSERT_FAILURE "(manufacture) != nullptr", "src/hci/manufacture.cpp:444", "updateProgressBar", "%s should not be nullptr", "manufacture": Exit Sub

    With aStruct(manufacture).functionality.factory
        If .psSubject <> 0 Then
            If .buildPointsRemaining < calcTemplateBuild(apsTemplates(.psSubject)) Then
                ' Started production. Set the colour of the bar to yellow.
                buildPointsTotal = calcTemplateBuild(apsTemplates(FactoryGetTemplate(manufacture)))
                If .timeStartHold = 0 Then buildRate = getBuildingProductionPoints(factory_) Else buildRate = 0
                formatTime progressBar, buildPointsTotal - .buildPointsRemaining, buildPointsTotal, buildRate, "Construction Progress"
                Exit Sub
            End If
        End If
        ' Not yet started production.
        neededPower = checkPowerRequest(factory_)
        If .psSubject <> 0 Then powerToBuild = calcTemplatePower(apsTemplates(.psSubject)) Else powerToBuild = 0
        formatPower progressBar, neededPower, powerToBuild
    End With
End Sub

Public Function ManufactureStatsButton_isHighlighted(ByVal self As WIDGET) As Boolean
    Dim controller As ManufactureController, factory_ As Long
    Set controller = ManufactureStatsButton_(self).controller
    factory_ = controller.getObjectAt(ManufactureStatsButton_(self).objectIndex)
    If factory_ <> 0 Then ManufactureStatsButton_isHighlighted = (aStruct(factory_).selected <> 0 Or factory_ = controller.getHighlightedObject())
End Function

Public Sub ManufactureStatsButton_clickPrimary(ByVal self As WIDGET)
    Dim controller As ManufactureController, factory_ As Long
    Set controller = ManufactureStatsButton_(self).controller
    factory_ = controller.getObjectAt(ManufactureStatsButton_(self).objectIndex)
    If Not (factory_ <> 0) Then ASSERT_FAILURE "(factory) != nullptr", "src/hci/manufacture.cpp:471", "clickPrimary", "%s should not be nullptr", "factory": Exit Sub
    controller.releaseFactoryProduction factory_
    BaseObjectsController_clearStructureSelection controller
    BaseObjectsController_selectObject controller, factory_, True
    BaseStatsController_scheduleDisplayStatsForm controller
End Sub

Public Sub ManufactureStatsButton_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean)
    Dim controller As ManufactureController, factory_ As Long
    Set controller = ManufactureStatsButton_(self).controller
    factory_ = controller.getObjectAt(ManufactureStatsButton_(self).objectIndex)
    If Not (factory_ <> 0) Then ASSERT_FAILURE "(factory) != nullptr", "src/hci/manufacture.cpp:481", "clickSecondary", "%s should not be nullptr", "factory": Exit Sub
    BaseObjectsController_clearStructureSelection controller
    controller.cancelFactoryProduction factory_
    controller.setHighlightedObject factory_, True
    controller.refresh
    BaseStatsController_scheduleDisplayStatsForm controller
End Sub

' MARK: - ManufactureOptionButton

Public Sub ManufactureOptionButton_construct(ByVal self As WIDGET)
    StatsFormButton_construct self
    self.kind = K_ManufactureOptionButton
End Sub

Public Function ManufactureOptionButton_make(ByVal controller As ManufactureController, ByVal manufactureOptionIndex As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_ManufactureOptionButton)
    ManufactureOptionButton_construct widget
    Set ManufactureOptionButton_(widget).controller = controller
    ManufactureOptionButton_(widget).manufactureOptionIndex = manufactureOptionIndex
    ManufactureOptionButton_initialize widget
    Set ManufactureOptionButton_make = widget
End Function

Public Sub ManufactureOptionButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim stat As Long
    vt_DynamicIntFancyButton_updateLayout self
    stat = vt_StatsButton_getStats(self)
    If Not (stat <> 0) Then ASSERT_FAILURE "(stat) != nullptr", "src/hci/manufacture.cpp:516", "display", "%s should not be nullptr", "stat": Exit Sub

    IntFancyButton_displayIMD self, AtlasImage_(), ImdObject_DroidTemplate(stat), xOffset, yOffset
    IntFancyButton_displayIfHighlight self, xOffset, yOffset
End Sub

' DROID_TEMPLATE *getStats() override: as the BASE_STATS * of the virtual
Public Function ManufactureOptionButton_getStats(ByVal self As WIDGET) As Long
    Dim psTemplate As Long
    psTemplate = ManufactureOptionButton_(self).controller.getStatsAt(ManufactureOptionButton_(self).manufactureOptionIndex)
    If psTemplate <> 0 Then ManufactureOptionButton_getStats = STAT_TEMPLATE + psTemplate
End Function

Public Sub ManufactureOptionButton_initialize(ByVal self As WIDGET)
    StatsFormButton_addCostBar self
    ManufactureOptionButton_addProductionRunSizeLabel self
End Sub

Public Sub ManufactureOptionButton_addProductionRunSizeLabel(ByVal self As WIDGET)
    Set ManufactureOptionButton_(self).productionRunSizeLabel = makeProductionRunSizeLabel()
    self.attach ManufactureOptionButton_(self).productionRunSizeLabel
End Sub

' droidTemplate: a DROID_TEMPLATE * (0 = nullptr)
Public Sub ManufactureOptionButton_updateProductionRunSizeLabel(ByVal self As WIDGET, ByVal structure_ As Long, ByVal droidTemplate As Long)
    Dim production As ProductionRunEntry, productionLoops As Long, labelText As String
    production = getProduction(structure_, droidTemplate)
    If ProductionRunEntry_isValid(production) Then
        productionLoops = getProductionLoops(structure_)
        If productionLoops > 0 Then
            labelText = va("%d/%d", ProductionRunEntry_numRemaining(production), production.quantity)
        Else
            labelText = va("%d", ProductionRunEntry_numRemaining(production), production.quantity)
        End If
        ManufactureOptionButton_(self).productionRunSizeLabel.setString labelText
        ManufactureOptionButton_(self).productionRunSizeLabel.show
    Else
        ManufactureOptionButton_(self).productionRunSizeLabel.hide
    End If
End Sub

Public Function ManufactureOptionButton_isHighlighted(ByVal self As WIDGET) As Boolean
    ManufactureOptionButton_isHighlighted = BaseObjectsStatsController_isHighlightedObjectStats(ManufactureOptionButton_(self).controller, ManufactureOptionButton_(self).manufactureOptionIndex)
End Function

Public Sub ManufactureOptionButton_updateLayout(ByVal self As WIDGET)
    DynamicIntFancyButton_updateLayout self
    If self.isMouseOverWidget() Then
        intSetShadowPower vt_StatsFormButton_getCost(self)
    End If
    W_BARGRAPH_(StatsFormButton_(self).costBar).majorSize = minL(100, vt_StatsFormButton_getCost(self) \ POWERPOINTS_DROIDDIV)
    ManufactureOptionButton_updateProductionRunSizeLabel self, ManufactureOptionButton_(self).controller.getHighlightedObject(), castDroidTemplate(vt_StatsButton_getStats(self))
End Sub

Public Function ManufactureOptionButton_getCost(ByVal self As WIDGET) As Long
    Dim psTemplate As Long
    psTemplate = castDroidTemplate(vt_StatsButton_getStats(self))
    If psTemplate <> 0 Then ManufactureOptionButton_getCost = calcTemplatePower(apsTemplates(psTemplate)) Else ManufactureOptionButton_getCost = 0
End Function

Public Sub ManufactureOptionButton_clickPrimary(ByVal self As WIDGET)
    ManufactureOptionButton_adjustFactoryProduction self, True
End Sub

Public Sub ManufactureOptionButton_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean)
    ManufactureOptionButton_adjustFactoryProduction self, False
End Sub

Public Sub ManufactureOptionButton_adjustFactoryProduction(ByVal self As WIDGET, ByVal add As Boolean)
    Dim clickedStats As Long, controllerRef As ManufactureController
    clickedStats = ManufactureOptionButton_(self).controller.getStatsAt(ManufactureOptionButton_(self).manufactureOptionIndex)
    If Not (clickedStats <> 0) Then ASSERT_FAILURE "(clickedStats) != nullptr", "src/hci/manufacture.cpp:590", "adjustFactoryProduction", "%s should not be nullptr", "clickedStats": Exit Sub

    Set controllerRef = ManufactureOptionButton_(self).controller
    widgScheduleTask WzFunc_new(FN_ManufactureOptionButton_adjust, add, clickedStats, controllerRef)
End Sub

' MARK: - ManufactureObjectsForm

Public Sub ManufactureObjectsForm_construct(ByVal self As WIDGET)
    ObjectsForm_construct self
    self.kind = K_ManufactureObjectsForm
End Sub

Public Function ManufactureObjectsForm_make(ByVal controller As ManufactureController) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_ManufactureObjectsForm)
    ManufactureObjectsForm_construct widget
    Set ManufactureObjectsForm_(widget).controller = controller
    ObjectsForm_initialize widget
    Set ManufactureObjectsForm_make = widget
End Function

Public Function ManufactureObjectsForm_makeStatsButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set ManufactureObjectsForm_makeStatsButton = ManufactureStatsButton_make(ManufactureObjectsForm_(self).controller, buttonIndex)
End Function

Public Function ManufactureObjectsForm_makeObjectButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set ManufactureObjectsForm_makeObjectButton = ManufactureObjectButton_make(ManufactureObjectsForm_(self).controller, buttonIndex)
End Function

Public Function ManufactureObjectsForm_getController(ByVal self As WIDGET) As Object
    Set ManufactureObjectsForm_getController = ManufactureObjectsForm_(self).controller
End Function

' MARK: - ManufactureStatsForm

Public Sub ManufactureStatsForm_construct(ByVal self As WIDGET)
    ObjectStatsForm_construct self
    self.kind = K_ManufactureStatsForm
End Sub

Public Function ManufactureStatsForm_make(ByVal controller As ManufactureController) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_ManufactureStatsForm)
    ManufactureStatsForm_construct widget
    Set ManufactureStatsForm_(widget).controller = controller
    vt_StatsForm_initialize widget
    Set ManufactureStatsForm_make = widget
End Function

Public Function ManufactureStatsForm_makeOptionButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set ManufactureStatsForm_makeOptionButton = ManufactureOptionButton_make(ManufactureStatsForm_(self).controller, buttonIndex)
End Function

Public Function ManufactureStatsForm_getController(ByVal self As WIDGET) As Object
    Set ManufactureStatsForm_getController = ManufactureStatsForm_(self).controller
End Function

Public Sub ManufactureStatsForm_initialize(ByVal self As WIDGET)
    StatsForm_initialize self                       ' BaseWidget::initialize(): ObjectStatsForm's is StatsForm's
    ManufactureStatsForm_addObsoleteButton self
    self.attach DeliveryPointButton_new(ManufactureStatsForm_(self).controller)
    ManufactureStatsForm_addLoopProductionWidgets self
End Sub

Public Sub ManufactureStatsForm_addObsoleteButton(ByVal self As WIDGET)
    Dim obsoleteButton As WIDGET, weakController As ManufactureController
    Set obsoleteButton = MultipleChoiceButton_new()
    self.attach obsoleteButton
    obsoleteButton.style = obsoleteButton.style Or WBUT_SECONDARY
    MultipleChoiceButton_setChoice obsoleteButton, IIf(ManufactureStatsForm_(self).controller.shouldShowRedundantDesign(), 1, 0)
    MultipleChoiceButton_setImages obsoleteButton, 0, W_BUTTON_Images_(AtlasImage_(IntImages, IMAGE_OBSOLETE_HIDE_UP), AtlasImage_(IntImages, IMAGE_OBSOLETE_HIDE_UP), AtlasImage_(IntImages, IMAGE_OBSOLETE_HIDE_HI), AtlasImage_())
    MultipleChoiceButton_setTip_choice obsoleteButton, 0, "Hiding Obsolete Tech"
    MultipleChoiceButton_setImages obsoleteButton, 1, W_BUTTON_Images_(AtlasImage_(IntImages, IMAGE_OBSOLETE_SHOW_UP), AtlasImage_(IntImages, IMAGE_OBSOLETE_SHOW_UP), AtlasImage_(IntImages, IMAGE_OBSOLETE_SHOW_HI), AtlasImage_())
    MultipleChoiceButton_setTip_choice obsoleteButton, 1, "Showing Obsolete Tech"
    obsoleteButton.move 4 + AtlasImage_width(IntImages, IMAGE_FDP_UP) + 4, STAT_SLDY

    Set weakController = ManufactureStatsForm_(self).controller ' std::weak_ptr<ManufactureController>(controller): a strong reference
    W_BUTTON_addOnClickHandler obsoleteButton, WzFunc_new(FN_ManufactureStatsForm_obsoleteClicked, weakController)
End Sub

Public Sub ManufactureStatsForm_addLoopProductionWidgets(ByVal self As WIDGET)
    Dim loopProductionButton As WIDGET, loopProductionLabel As WIDGET
    Set loopProductionButton = LoopProductionButton_new(ManufactureStatsForm_(self).controller)
    self.attach loopProductionButton
    loopProductionButton.move STAT_SLDX + STAT_SLDWIDTH + 2, STAT_SLDY

    Set loopProductionLabel = LoopProductionLabel_new(ManufactureStatsForm_(self).controller)
    self.attach loopProductionLabel
    loopProductionLabel.setGeometry loopProductionButton.x() - 15, loopProductionButton.y(), 12, 15
End Sub

' MARK: - ManufactureStatsForm::DeliveryPointButton

Public Sub DeliveryPointButton_construct(ByVal self As WIDGET, ByVal controller As ManufactureController)
    Dim weakController As ManufactureController
    W_BUTTON_construct self
    self.kind = K_DeliveryPointButton
    Set DeliveryPointButton_(self).controller = controller
    self.style = self.style Or WBUT_SECONDARY
    self.move 4, STAT_SLDY
    self.setTip "Factory Delivery Point"

    Set weakController = controller     ' std::weak_ptr<ManufactureController>(controller): a strong reference
    W_BUTTON_addOnClickHandler self, WzFunc_new(FN_DeliveryPointButton_clicked, weakController)
End Sub

Public Function DeliveryPointButton_new(ByVal controller As ManufactureController) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_DeliveryPointButton)
    DeliveryPointButton_construct widget, controller
    Set DeliveryPointButton_new = widget
End Function

Public Sub DeliveryPointButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    DeliveryPointButton_updateLayout self
    W_BUTTON_display self, xOffset, yOffset
End Sub

Public Sub DeliveryPointButton_updateLayout(ByVal self As WIDGET)
    Dim factory_ As Long
    factory_ = DeliveryPointButton_(self).controller.getHighlightedObject()
    If Not (factory_ <> 0) Then ASSERT_FAILURE "(factory) != nullptr", "src/hci/manufacture.cpp:743", "updateLayout", "%s should not be nullptr", "factory": Exit Sub

    Select Case asStructureStats(aStruct(factory_).pStructureType).typ
    Case REF_CYBORG_FACTORY
        W_BUTTON_setImages_atlas self, AtlasImage_(IntImages, IMAGE_CDP_UP), AtlasImage_(IntImages, IMAGE_CDP_DOWN), AtlasImage_(IntImages, IMAGE_CDP_HI), AtlasImage_()
    Case REF_VTOL_FACTORY
        W_BUTTON_setImages_atlas self, AtlasImage_(IntImages, IMAGE_VDP_UP), AtlasImage_(IntImages, IMAGE_VDP_DOWN), AtlasImage_(IntImages, IMAGE_VDP_HI), AtlasImage_()
    Case Else                           ' default: case REF_FACTORY:
        W_BUTTON_setImages_atlas self, AtlasImage_(IntImages, IMAGE_FDP_UP), AtlasImage_(IntImages, IMAGE_FDP_DOWN), AtlasImage_(IntImages, IMAGE_FDP_HI), AtlasImage_()
    End Select
End Sub

' MARK: - ManufactureStatsForm::LoopProductionButton

Public Sub LoopProductionButton_construct(ByVal self As WIDGET, ByVal controller As ManufactureController)
    W_BUTTON_construct self
    self.kind = K_LoopProductionButton
    Set LoopProductionButton_(self).controller = controller
    self.style = self.style Or WBUT_SECONDARY
    W_BUTTON_setImages_atlas self, AtlasImage_(IntImages, IMAGE_LOOP_UP), AtlasImage_(IntImages, IMAGE_LOOP_DOWN), AtlasImage_(IntImages, IMAGE_LOOP_HI), AtlasImage_()
    self.setTip "Loop Production"
End Sub

Public Function LoopProductionButton_new(ByVal controller As ManufactureController) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_LoopProductionButton)
    LoopProductionButton_construct widget, controller
    Set LoopProductionButton_new = widget
End Function

Public Sub LoopProductionButton_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    W_BUTTON_released self, psContext, key
    LoopProductionButton_(self).controller.adjustFactoryLoop (key = WKEY_PRIMARY)
End Sub

Public Sub LoopProductionButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    LoopProductionButton_updateLayout self
    W_BUTTON_display self, xOffset, yOffset
End Sub

Public Sub LoopProductionButton_updateLayout(ByVal self As WIDGET)
    If getProductionLoops(LoopProductionButton_(self).controller.getHighlightedObject()) = 0 Then
        self.setState 0
    Else
        self.setState WBUT_CLICKLOCK
    End If
End Sub

' MARK: - ManufactureStatsForm::LoopProductionLabel

Public Sub LoopProductionLabel_construct(ByVal self As WIDGET, ByVal controller As ManufactureController)
    W_LABEL_construct self
    self.kind = K_LoopProductionLabel
    LoopProductionLabel_(self).lastProductionLoop = 0
    Set LoopProductionLabel_(self).controller = controller
End Sub

Public Function LoopProductionLabel_new(ByVal controller As ManufactureController) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_LoopProductionLabel)
    LoopProductionLabel_construct widget, controller
    Set LoopProductionLabel_new = widget
End Function

Public Sub LoopProductionLabel_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    LoopProductionLabel_updateLayout self
    W_LABEL_display self, xOffset, yOffset
End Sub

Public Sub LoopProductionLabel_updateLayout(ByVal self As WIDGET)
    Dim productionLoops As Long
    productionLoops = getProductionLoops(LoopProductionLabel_(self).controller.getHighlightedObject())
    If productionLoops <> LoopProductionLabel_(self).lastProductionLoop Then
        LoopProductionLabel_(self).lastProductionLoop = productionLoops
        self.setString LoopProductionLabel_getNewString(self)
    End If
End Sub

Public Function LoopProductionLabel_getNewString(ByVal self As WIDGET) As String
    Select Case LoopProductionLabel_(self).lastProductionLoop
    Case 0
        LoopProductionLabel_getNewString = ""
    Case INFINITE_PRODUCTION
        LoopProductionLabel_getNewString = ChrW(&H221E)  ' the infinity sign, U+221E
    Case Else
        LoopProductionLabel_getNewString = va("%u", LoopProductionLabel_(self).lastProductionLoop + 1)
    End Select
End Function

' MARK: - the lambdas

Public Function manufacture_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim controllerRef As ManufactureController, manufactureController As ManufactureController, button As WIDGET, newValue As Boolean
    Select Case f.id
    Case FN_ManufactureOptionButton_adjust
        Set controllerRef = f.cap(2)
        controllerRef.adjustFactoryProduction f.cap(1), f.cap(0)
    Case FN_ManufactureStatsForm_obsoleteClicked
        Set manufactureController = f.cap(0) ' weakController.lock()
        If Not manufactureController Is Nothing Then
            Set button = args(0)        ' auto &obsoleteButton = static_cast<MultipleChoiceButton &>(button);
            newValue = Not (MultipleChoiceButton_getChoice(button) <> 0)
            manufactureController.setShouldShowRedundantDesign newValue
            MultipleChoiceButton_setChoice button, IIf(newValue, 1, 0)
        End If
    Case FN_DeliveryPointButton_clicked
        Set manufactureController = f.cap(0) ' auto controller = weakController.lock();
        If Not (Not manufactureController Is Nothing) Then ASSERT_FAILURE "(controller) != nullptr", "src/hci/manufacture.cpp:727", "DeliveryPointButton", "%s should not be nullptr", "controller": Exit Function
        manufactureController.startDeliveryPointPosition
    End Select
End Function
