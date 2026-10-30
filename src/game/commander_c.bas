Attribute VB_Name = "commander_c"
' src/hci/commander.cpp, commander.h: the commander panel's widgets and CommanderController's
' static member (the controller itself is the class CommanderController).
' (VB) widget class CommanderObjectButton : ObjectButton
' (VB) widget class CommanderStatsButton : StatsButton
' (VB) widget class CommanderObjectsForm : ObjectsForm
Option Explicit

Public CommanderController_highlightedCommander As Long ' static DROID *CommanderController::highlightedCommander = nullptr

Public Const FN_CommanderController_displayOrderForm_task As Long = 8400 ' (cap 0: psWeakControllerRef)

' MARK: - CommanderObjectButton

Public Sub CommanderObjectButton_construct(ByVal self As WIDGET)
    ObjectButton_construct self
    self.kind = K_CommanderObjectButton
End Sub

Public Function CommanderObjectButton_make(ByVal controller As CommanderController, ByVal objectIndex As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_CommanderObjectButton)
    CommanderObjectButton_construct widget
    Set CommanderObjectButton_(widget).controller = controller
    ObjectButton_(widget).objectIndex = objectIndex
    CommanderObjectButton_initialize widget
    Set CommanderObjectButton_make = widget
End Function

Public Sub CommanderObjectButton_clickPrimary(ByVal self As WIDGET)
    Dim controller As CommanderController
    Set controller = CommanderObjectButton_(self).controller
    BaseObjectsController_clearSelection controller
    BaseObjectsController_selectObject controller, controller.getObjectAt(ObjectButton_(self).objectIndex), False
    vt_ObjectButton_jump self
    controller.displayOrderForm
End Sub

Public Sub CommanderObjectButton_initialize(ByVal self As WIDGET)
    With CommanderObjectButton_(self)
        Set .groupSizeLabel = W_LABEL_new()
        self.attach .groupSizeLabel
        .groupSizeLabel.setGeometry OBJ_TEXTX, OBJ_B1TEXTY, 16, 16

        Set .experienceStarsLabel = W_LABEL_new()
        self.attach .experienceStarsLabel
        .experienceStarsLabel.setGeometry STAT_POWERBARX, STAT_POWERBARY, 16, 16
    End With
End Sub

Public Sub CommanderObjectButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim droid_ As Long
    vt_DynamicIntFancyButton_updateLayout self
    droid_ = CommanderObjectButton_(self).controller.getObjectAt(ObjectButton_(self).objectIndex)
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/commander.cpp:144", "display", "%s should not be nullptr", "droid": Exit Sub
    If isDead(droid_) Then
        ASSERT_FAILURE "!isDead(droid)", "src/hci/commander.cpp:147", "display", "Droid is dead"
        ' ensure the backing information is refreshed before the next draw
        intRefreshScreen
        Exit Sub
    End If
    IntFancyButton_displayIMD self, AtlasImage_(), ImdObject_Droid(droid_), xOffset, yOffset
    IntFancyButton_displayIfHighlight self, xOffset, yOffset
End Sub

Public Sub CommanderObjectButton_updateLayout(ByVal self As WIDGET)
    Dim droid_ As Long
    DynamicIntFancyButton_updateLayout self
    droid_ = CommanderObjectButton_(self).controller.getObjectAt(ObjectButton_(self).objectIndex)
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/commander.cpp:160", "updateLayout", "%s should not be nullptr", "droid": Exit Sub
    CommanderObjectButton_updateGroupSizeLabel self, droid_
    CommanderObjectButton_updateExperienceStarsLabel self, droid_
End Sub

Public Sub CommanderObjectButton_updateGroupSizeLabel(ByVal self As WIDGET, ByVal droid_ As Long)
    Dim text As String, members As Long
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/commander.cpp:167", "updateGroupSizeLabel", "%s should not be nullptr", "droid": Exit Sub
    If aDroid(droid_).psGroup <> 0 Then members = DROID_GROUP_getNumMembers(aDroid(droid_).psGroup) Else members = 0
    text = va("%u/%u", members, cmdDroidMaxGroup(droid_))
    CommanderObjectButton_(self).groupSizeLabel.setString text
    CommanderObjectButton_(self).groupSizeLabel.show
End Sub

Public Sub CommanderObjectButton_updateExperienceStarsLabel(ByVal self As WIDGET, ByVal droid_ As Long)
    Dim numStars As Long
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/commander.cpp:175", "updateExperienceStarsLabel", "%s should not be nullptr", "droid": Exit Sub
    numStars = getDroidLevel(droid_) - 1
    If numStars < 0 Then numStars = 0   ' std::max((int)getDroidLevel(droid) - 1, 0)
    CommanderObjectButton_(self).experienceStarsLabel.setString String$(numStars, "*")
    CommanderObjectButton_(self).experienceStarsLabel.show
End Sub

Public Function CommanderObjectButton_getController(ByVal self As WIDGET) As Object
    Set CommanderObjectButton_getController = CommanderObjectButton_(self).controller
End Function

Public Function CommanderObjectButton_getTip(ByVal self As WIDGET) As String
    Dim droid_ As Long
    droid_ = CommanderObjectButton_(self).controller.getObjectAt(ObjectButton_(self).objectIndex)
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/commander.cpp:189", "getTip", "%s should not be nullptr", "droid": CommanderObjectButton_getTip = "": Exit Function
    CommanderObjectButton_getTip = droidGetName(droid_)
End Function

' MARK: - CommanderStatsButton

Public Sub CommanderStatsButton_construct(ByVal self As WIDGET)
    StatsButton_construct self
    self.kind = K_CommanderStatsButton
End Sub

Public Function CommanderStatsButton_make(ByVal controller As CommanderController, ByVal objectIndex As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_CommanderStatsButton)
    CommanderStatsButton_construct widget
    Set CommanderStatsButton_(widget).controller = controller
    CommanderStatsButton_(widget).objectIndex = objectIndex
    CommanderStatsButton_initialize widget
    Set CommanderStatsButton_make = widget
End Function

Public Sub CommanderStatsButton_initialize(ByVal self As WIDGET)
    With CommanderStatsButton_(self)
        Set .assignedFactoriesLabel = W_LABEL_new()
        self.attach .assignedFactoriesLabel
        .assignedFactoriesLabel.setGeometry OBJ_TEXTX, OBJ_T1TEXTY, 16, 16

        Set .assignedCyborgFactoriesLabel = W_LABEL_new()
        self.attach .assignedCyborgFactoriesLabel
        .assignedCyborgFactoriesLabel.setGeometry OBJ_TEXTX, OBJ_T2TEXTY, 16, 16

        Set .assignedVtolFactoriesLabel = W_LABEL_new()
        self.attach .assignedVtolFactoriesLabel
        .assignedVtolFactoriesLabel.setGeometry OBJ_TEXTX, OBJ_T3TEXTY, 16, 16
    End With
End Sub

Public Sub CommanderStatsButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim factory_ As Long
    vt_DynamicIntFancyButton_updateLayout self
    factory_ = CommanderStatsButton_(self).controller.getAssignedFactoryAt(CommanderStatsButton_(self).objectIndex)
    If factory_ <> 0 Then
        IntFancyButton_displayIMD self, AtlasImage_(), ImdObject_Structure(factory_), xOffset, yOffset
    Else
        IntFancyButton_displayIMD self, AtlasImage_(), ImdObject_Component(0), xOffset, yOffset
    End If
    IntFancyButton_displayIfHighlight self, xOffset, yOffset
End Sub

Public Sub CommanderStatsButton_updateLayout(ByVal self As WIDGET)
    Dim droid_ As Long
    DynamicIntFancyButton_updateLayout self
    droid_ = CommanderStatsButton_(self).controller.getObjectAt(CommanderStatsButton_(self).objectIndex)
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/commander.cpp:244", "updateLayout", "%s should not be nullptr", "droid": Exit Sub
    CommanderStatsButton_updateAssignedFactoriesLabel self, CommanderStatsButton_(self).assignedFactoriesLabel, droid_, DSS_ASSPROD_SHIFT
    CommanderStatsButton_updateAssignedFactoriesLabel self, CommanderStatsButton_(self).assignedCyborgFactoriesLabel, droid_, DSS_ASSPROD_CYBORG_SHIFT
    CommanderStatsButton_updateAssignedFactoriesLabel self, CommanderStatsButton_(self).assignedVtolFactoriesLabel, droid_, DSS_ASSPROD_VTOL_SHIFT
End Sub

Public Sub CommanderStatsButton_updateAssignedFactoriesLabel(ByVal self As WIDGET, ByVal label As WIDGET, ByVal droid_ As Long, ByVal factoryTypeShift As Long)
    Const maxAssignedFactories As Long = 5
    Dim text As String, index As Long, i As Long
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/commander.cpp:253", "updateAssignedFactoriesLabel", "%s should not be nullptr", "droid": Exit Sub
    '/**
    ' * TODO Support up to MAX_FACTORY (which won't fit in the ugly secondaryOrder bitmask hack).
    ' * Comment taken from commit 34d8148e4a
    ' **/
    ' char text[maxAssignedFactories + 1] = {0};

    index = 0
    For i = 0 To maxAssignedFactories - 1
        If (aDroid(droid_).secondaryOrder And SHL(1, i + factoryTypeShift)) <> 0 Then
            text = text & Chr$(Asc("1") + i)
            index = index + 1
        End If
    Next

    If index > 0 Then
        label.setString text
        label.show
    Else
        label.hide
    End If
End Sub

' STRUCTURE_STATS *getStats() override: as the BASE_STATS * of the virtual
Public Function CommanderStatsButton_getStats(ByVal self As WIDGET) As Long
    CommanderStatsButton_getStats = structureStatsBase(CommanderStatsButton_(self).controller.getObjectStatsAt(CommanderStatsButton_(self).objectIndex))
End Function

Public Function CommanderStatsButton_isHighlighted(ByVal self As WIDGET) As Boolean
    Dim controller As CommanderController, droid_ As Long
    Set controller = CommanderStatsButton_(self).controller
    droid_ = controller.getObjectAt(CommanderStatsButton_(self).objectIndex)
    If droid_ <> 0 Then CommanderStatsButton_isHighlighted = (droid_ = controller.getHighlightedObject())
End Function

Public Sub CommanderStatsButton_clickPrimary(ByVal self As WIDGET)
    Dim controller As CommanderController, droid_ As Long
    Set controller = CommanderStatsButton_(self).controller
    droid_ = controller.getObjectAt(CommanderStatsButton_(self).objectIndex)
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/commander.cpp:295", "clickPrimary", "%s should not be nullptr", "droid": Exit Sub
    BaseObjectsController_clearSelection controller
    BaseObjectsController_selectObject controller, droid_, True
    controller.displayOrderForm
End Sub

Public Sub CommanderStatsButton_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean)
    Dim controller As CommanderController, droid_ As Long, highlighted As Long, ok As Boolean
    Set controller = CommanderStatsButton_(self).controller
    droid_ = controller.getObjectAt(CommanderStatsButton_(self).objectIndex)
    If Not (droid_ <> 0) Then ASSERT_FAILURE "(droid) != nullptr", "src/hci/commander.cpp:304", "clickSecondary", "%s should not be nullptr", "droid": Exit Sub
    highlighted = controller.getHighlightedObject()

    ' prevent highlighting a commander when another commander is already selected
    If droid_ = highlighted Then
        ok = True
    ElseIf highlighted <> 0 Then
        ok = aDroid(highlighted).selected = 0
    End If
    If ok Then
        controller.setHighlightedObject droid_, True
        controller.displayOrderForm
    End If
End Sub

' MARK: - CommanderObjectsForm

Public Sub CommanderObjectsForm_construct(ByVal self As WIDGET)
    ObjectsForm_construct self
    self.kind = K_CommanderObjectsForm
End Sub

Public Function CommanderObjectsForm_make(ByVal controller As CommanderController) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_CommanderObjectsForm)
    CommanderObjectsForm_construct widget
    Set CommanderObjectsForm_(widget).controller = controller
    ObjectsForm_initialize widget
    Set CommanderObjectsForm_make = widget
End Function

Public Function CommanderObjectsForm_makeStatsButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set CommanderObjectsForm_makeStatsButton = CommanderStatsButton_make(CommanderObjectsForm_(self).controller, buttonIndex)
End Function

Public Function CommanderObjectsForm_makeObjectButton(ByVal self As WIDGET, ByVal buttonIndex As Long) As WIDGET
    Set CommanderObjectsForm_makeObjectButton = CommanderObjectButton_make(CommanderObjectsForm_(self).controller, buttonIndex)
End Function

Public Function CommanderObjectsForm_getController(ByVal self As WIDGET) As Object
    Set CommanderObjectsForm_getController = CommanderObjectsForm_(self).controller
End Function

' MARK: - the lambdas

Public Function commander_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psDroid As Long, strongControllerRef As CommanderController
    Select Case f.id
    Case FN_CommanderController_displayOrderForm_task
        psDroid = 0
        ' (VB) psWeakControllerRef.lock(): the capture is a strong reference; the controller is alive
        ' while its objects form (which holds it) is on the screen
        If Not widgGetFromID(psWScreen, IDOBJ_FORM) Is Nothing Then
            Set strongControllerRef = f.cap(0)
            psDroid = strongControllerRef.getHighlightedObject()
        End If
        intAddOrder psDroid
        intMode = INT_CMDORDER
    End Select
End Function
