Attribute VB_Name = "groups_c"
' src/hci/groups.cpp, groups.h: the unit groups bar (GroupsForum, its GroupButtons) and the group
' button switch. GroupsUIController and its GroupDisplayInfo are classes of those names.
' (VB) widget class GroupsForum : IntFormAnimated
' (VB) widget class GroupButton : DynamicIntFancyButton
Option Explicit

Private groupButtonEnabled As Boolean   ' static bool groupButtonEnabled = true (groups_c_init)

Private Const MAX_DAMAGE_GLOW_ALPHA As Long = 100

Public Const FN_GroupsForum_calcLayout As Long = 8500

' (VB) the C++ initialisers of this module's globals (Sub Main calls it)
Public Sub groups_c_init()
    groupButtonEnabled = True
End Sub

Public Sub setGroupButtonEnabled(ByVal bNewState As Boolean)
    groupButtonEnabled = bNewState
End Sub

Public Function getGroupButtonEnabled() As Boolean
    getGroupButtonEnabled = groupButtonEnabled
End Function

' MARK: - GroupButton

Public Sub GroupButton_construct(ByVal self As WIDGET)
    DynamicIntFancyButton_construct self
    self.kind = K_GroupButton
    GroupButton_(self).lastUpdatedGlowAlphaTime = 0
End Sub

Public Function GroupButton_make(ByVal controller As GroupsUIController, ByVal groupNumber As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_GroupButton)
    GroupButton_construct widget
    Set GroupButton_(widget).controller = controller
    GroupButton_(widget).groupNumber = groupNumber
    GroupButton_initialize widget
    Set GroupButton_make = widget
End Function

Public Sub GroupButton_initialize(ByVal self As WIDGET)
    Dim groupNumber As Long, groupNumberStr As String, helpInfo As WidgetHelp
    groupNumber = GroupButton_(self).groupNumber
    With GroupButton_(self)
        Set .groupNumberLabel = W_LABEL_new()
        self.attach .groupNumberLabel
        .groupNumberLabel.setGeometry OBJ_TEXTX, OBJ_B1TEXTY - 5, 16, 16
        .groupNumberLabel.setString va("%u", groupNumber)
        .groupNumberLabel.setTransparentToMouse True

        Set .groupCountLabel = W_LABEL_new()
        self.attach .groupCountLabel
        .groupCountLabel.setGeometry OBJ_TEXTX + 40, OBJ_B1TEXTY + 20, 16, 16
        .groupCountLabel.setString ""
        .groupCountLabel.setTransparentToMouse True

        Set .groupDamagedCountLabel = W_LABEL_new()
        self.attach .groupDamagedCountLabel
        .groupDamagedCountLabel.setGeometry 0, 0, 16, 16
        W_LABEL_setFontColour .groupDamagedCountLabel, pal_RGBA(255, 0, 0, 255) ' red
        .groupDamagedCountLabel.setString ""
        .groupDamagedCountLabel.setTransparentToMouse True
    End With

    IntFancyButton_(self).buttonBackgroundEmpty = True

    groupNumberStr = CStr(groupNumber)  ' std::to_string(groupNumber)
    self.setTip "Select / Assign Group Number: " & groupNumberStr

    Set helpInfo = New WidgetHelp
    helpInfo.setTitle va("Group %u", groupNumber)
    helpInfo.addInteraction WidgetHelp_triggers(InteractionTriggers_PrimaryClick), va("Select the Units in Group %u", groupNumber)
    helpInfo.addInteraction WidgetHelp_triggers(InteractionTriggers_SecondaryClick, InteractionTriggers_ClickAndHold), va("Assign Selected Units to Group %u", groupNumber)
    helpInfo.addInteraction WidgetHelp_triggers(InteractionTriggers_Misc), "Center Camera on this Group by clicking or tapping twice"
    helpInfo.addRelatedKeybinding "SelectGrouping_" & groupNumberStr
    helpInfo.addRelatedKeybinding "AssignGrouping_" & groupNumberStr
    helpInfo.addRelatedKeybinding "AddGrouping_" & groupNumberStr

    self.setHelp helpInfo
End Sub

Public Sub GroupButton_clickPrimary(ByVal self As WIDGET)
    GroupButton_(self).controller.selectGroup GroupButton_(self).groupNumber
End Sub

Public Sub GroupButton_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean)
    GroupButton_(self).controller.assignSelectedDroidsToGroup GroupButton_(self).groupNumber
End Sub

Public Sub GroupButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim groupInfo As GroupDisplayInfo, targetAlpha As Long, fadeAmount As Long, xNumberOffset As Long
    Const xFitNumberInTheBox As Long = 16
    Dim groupCountLabel As WIDGET, groupDamagedCountLabel As WIDGET
    Set groupInfo = GroupButton_(self).controller.getGroupInfoAt(GroupButton_(self).groupNumber)
    If groupInfo Is Nothing Then
        Exit Sub
    End If
    Set groupCountLabel = GroupButton_(self).groupCountLabel
    Set groupDamagedCountLabel = GroupButton_(self).groupDamagedCountLabel

    If groupInfo.lastAccumulatedDamage > 0 Then
        targetAlpha = GroupButton_accumulatedDamageToTargetGlowAlpha(self, groupInfo.lastAccumulatedDamage, groupInfo.totalGroupMaxHealth, groupInfo.lastUnitsKilled)
        If targetAlpha > groupInfo.currAttackGlowAlpha Then groupInfo.currAttackGlowAlpha = targetAlpha
        GroupButton_(self).lastUpdatedGlowAlphaTime = realTime
        ' reset consumed values
        groupInfo.lastAccumulatedDamage = 0
        groupInfo.lastUnitsKilled = 0
    ElseIf groupInfo.currAttackGlowAlpha > 0 And ULT(100, realTime - GroupButton_(self).lastUpdatedGlowAlphaTime) Then
        fadeAmount = 50 * (realTime - GroupButton_(self).lastUpdatedGlowAlphaTime) \ GAME_TICKS_PER_SEC
        If fadeAmount > groupInfo.currAttackGlowAlpha Then
            groupInfo.currAttackGlowAlpha = 0
        Else
            groupInfo.currAttackGlowAlpha = groupInfo.currAttackGlowAlpha - (fadeAmount And &HFF&)
        End If
        GroupButton_(self).lastUpdatedGlowAlphaTime = realTime
    End If

    If groupInfo.numberInGroup = 0 And groupInfo.numberDamagedInGroup = 0 Then
        groupCountLabel.setString ""
        IntFancyButton_displayBlank self, xOffset, yOffset, False
    Else
        IntFancyButton_displayIMD self, AtlasImage_(), ImdObject_DroidTemplate(STAT_TEMPLATE + groupInfo.displayDroidTemplate), xOffset, yOffset
        groupCountLabel.setString va("%u", groupInfo.numberInGroup)
        xNumberOffset = 0
        If W_LABEL_getMaxLineWidth(groupCountLabel) > xFitNumberInTheBox Then
            xNumberOffset = xNumberOffset - (W_LABEL_getMaxLineWidth(groupCountLabel) - xFitNumberInTheBox)
        End If
        groupCountLabel.move OBJ_TEXTX + 40 + xNumberOffset - W_LABEL_getMaxLineWidth(groupDamagedCountLabel), groupCountLabel.y()
    End If

    If groupInfo.numberDamagedInGroup = 0 Then
        groupDamagedCountLabel.setString ""
    Else
        groupDamagedCountLabel.setString va("+%u", groupInfo.numberDamagedInGroup)
        groupDamagedCountLabel.move groupCountLabel.x() + W_LABEL_getMaxLineWidth(groupCountLabel), groupCountLabel.y()
    End If

    If groupInfo.currAttackGlowAlpha > 0 Then
        ' Based on damage, display an inner glow (animated)
        iV_DrawImageTint IntImages, IMAGE_BUT_INNER_GLOW, xOffset + self.x(), yOffset + self.y(), pal_RGBA(170, 0, 0, groupInfo.currAttackGlowAlpha)
    End If
End Sub

Public Function GroupButton_isHighlighted(ByVal self As WIDGET) As Boolean
    GroupButton_isHighlighted = False
End Function

' inline uint8_t accumulatedDamageToTargetGlowAlpha(uint64_t accumulatedDamage, uint64_t totalGroupMaxHealth, uint64_t unitsKilled)
Public Function GroupButton_accumulatedDamageToTargetGlowAlpha(ByVal self As WIDGET, ByVal accumulatedDamage As Double, ByVal totalGroupMaxHealth As Double, ByVal unitsKilled As Double) As Long
    Dim damageVisualPercent As Double
    If totalGroupMaxHealth > 0 Then damageVisualPercent = Fix((accumulatedDamage * 100) / totalGroupMaxHealth) Else damageVisualPercent = 0
    If unitsKilled > 0 Then
        damageVisualPercent = 100
    End If
    If damageVisualPercent > 100 Then damageVisualPercent = 100 ' std::min<uint64_t>(damageVisualPercent, 100)
    GroupButton_accumulatedDamageToTargetGlowAlpha = CLng(damageVisualPercent) * MAX_DAMAGE_GLOW_ALPHA \ 100
End Function

' MARK: - GroupsForum

Public Sub GroupsForum_construct(ByVal self As WIDGET)
    IntFormAnimated_construct self
    self.kind = K_GroupsForum
End Sub

Public Function GroupsForum_make() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_GroupsForum)
    GroupsForum_construct widget
    ' widget->controller = controller;
    GroupsForum_initialize widget
    Set GroupsForum_make = widget
End Function

Public Sub GroupsForum_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' draw the background
    IntFormAnimated_display self, xOffset, yOffset
End Sub

Public Sub GroupsForum_initialize(ByVal self As WIDGET)
    Dim i As Long, buttonHolder As WIDGET, groupButton As WIDGET, descriptionStr As String, help As WidgetHelp
    Set GroupsForum_(self).groupsUIController = New GroupsUIController

    ' the layout should be like this when the build menu is open
    self.id = IDOBJ_GROUP
    self.setCalcLayout WzFunc_new(FN_GroupsForum_calcLayout)
    GroupsForum_addTabList self
    ' create the 11 buttons for each group
    For i = 1 To 10
        ' check if the 10th group works
        Set buttonHolder = WIDGET_new()
        ListTabWidget_addWidgetToLayout GroupsForum_(self).groupsList, buttonHolder
        Set groupButton = GroupsForum_makeGroupButton(self, i Mod 10)
        buttonHolder.attach groupButton
        groupButton.setGeometry 0, 0, OBJ_BUTWIDTH, OBJ_BUTHEIGHT
    Next

    descriptionStr = "View and configure groups of units, which can be quickly selected and ordered."
    descriptionStr = descriptionStr & vbLf & vbLf
    descriptionStr = descriptionStr & "Group buttons will glow red when units are lost (or taking lots of damage)."
    Set help = New WidgetHelp
    help.setTitle "Unit Groups"
    help.setDescription descriptionStr
    help.addInteraction WidgetHelp_triggers(InteractionTriggers_PrimaryClick), "Select a Group"
    help.addInteraction WidgetHelp_triggers(InteractionTriggers_SecondaryClick, InteractionTriggers_ClickAndHold), "Assign Selected Units to a Group"
    help.addInteraction WidgetHelp_triggers(InteractionTriggers_Misc), "Center Camera on a Group by clicking or tapping twice on the group button"
    self.setHelp help
End Sub

Public Sub GroupsForum_updateData(ByVal self As WIDGET)
    GroupsForum_(self).groupsUIController.updateData
End Sub

Public Sub GroupsForum_updateSelectedGroup(ByVal self As WIDGET, ByVal group As Long)
    Dim groupButtonIdx As Long
    If group > 0 Then groupButtonIdx = group - 1 Else groupButtonIdx = 9
    If groupButtonIdx >= ListTabWidget_childrenSize(GroupsForum_(self).groupsList) Then
        Exit Sub
    End If
    ' ensure that the page that contains this group is displayed
    ListTabWidget_goToChildPage GroupsForum_(self).groupsList, groupButtonIdx
End Sub

Public Sub GroupsForum_addGroupDamageForCurrentTick(ByVal self As WIDGET, ByVal group As Long, ByVal additionalDamage As Double, ByVal unitKilled As Boolean)
    GroupsForum_(self).groupsUIController.addGroupDamageForCurrentTick group, additionalDamage, unitKilled
End Sub

' psCommander: a const DROID *
Public Sub GroupsForum_addCommanderGroupDamageForCurrentTick(ByVal self As WIDGET, ByVal psCommander As Long, ByVal additionalDamage As Double, ByVal unitKilled As Boolean)
    GroupsForum_(self).groupsUIController.addCommanderGroupDamageForCurrentTick psCommander, additionalDamage, unitKilled
End Sub

' WidgetHelp const * getHelp() const override: Nothing = nullptr
Public Function GroupsForum_getHelp(ByVal self As WIDGET) As WidgetHelp
    Set GroupsForum_getHelp = GroupsForum_(self).help
End Function

' optional<WidgetHelp>: Nothing = nullopt
Public Sub GroupsForum_setHelp(ByVal self As WIDGET, ByVal help As WidgetHelp)
    Set GroupsForum_(self).help = help
End Sub

Public Sub GroupsForum_addTabList(ByVal self As WIDGET)
    Dim groupsList As WIDGET, groupListWidth As Long, unitGroupsStr As String
    Set groupsList = IntListTabWidget_make(TabAlignment_RightAligned)
    Set GroupsForum_(self).groupsList = groupsList
    self.attach groupsList
    ListTabWidget_setChildSize groupsList, OBJ_BUTWIDTH, OBJ_BUTHEIGHT * 2
    ListTabWidget_setChildSpacing groupsList, OBJ_GAP, OBJ_GAP
    groupListWidth = OBJ_BUTWIDTH * 5 + STAT_GAP * 4
    groupsList.setGeometry (OBJ_BACKWIDTH - groupListWidth) \ 2, OBJ_TABY, groupListWidth, OBJ_BACKHEIGHT - OBJ_TABY
    unitGroupsStr = "Unit Groups:"
    unitGroupsStr = unitGroupsStr & " "
    ListTabWidget_setTitle groupsList, unitGroupsStr
End Sub

Public Function GroupsForum_makeGroupButton(ByVal self As WIDGET, ByVal groupNumber As Long) As WIDGET
    Set GroupsForum_makeGroupButton = GroupButton_make(GroupsForum_(self).groupsUIController, groupNumber)
End Function

' MARK: - the lambdas

Public Function groups_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET
    Select Case f.id
    Case FN_GroupsForum_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry GROUP_BACKX, GROUP_BACKY, GROUP_BACKWIDTH, GROUP_BACKHEIGHT
    End Select
End Function
