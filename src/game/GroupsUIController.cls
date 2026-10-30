VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GroupsUIController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/hci/groups.cpp: class GroupsUIController : public std::enable_shared_from_this<GroupsUIController>
' (not a widget). std::array<GroupDisplayInfo, 10> groupInfo is ten GroupDisplayInfo objects (the
' class of GroupsUIController::GroupDisplayInfo); a GroupDisplayInfo * is the object (Nothing =
' nullptr). A uint64_t is a Double.
Option Explicit

Private groupInfo(0 To 9) As GroupDisplayInfo

Private Sub Class_Initialize()
    Dim i As Long
    For i = 0 To 9
        Set groupInfo(i) = New GroupDisplayInfo
    Next
End Sub

Public Function getGroupInfoAt(ByVal index As Long) As GroupDisplayInfo
    If Not (index < 10) Then ASSERT_FAILURE "index < groupInfo.size()", "src/hci/groups.cpp:58", "getGroupInfoAt", "Invalid group index (%zu); max: (%zu)", index, 10: Set getGroupInfoAt = Nothing: Exit Function
    Set getGroupInfoAt = groupInfo(index)
End Function

Public Function size() As Long
    size = 10
End Function

Public Sub addGroupDamageForCurrentTick(ByVal group As Long, ByVal additionalDamage As Double, ByVal unitKilled As Boolean)
    Dim curr As GroupDisplayInfo
    Set curr = groupInfo(group)
    curr.lastAttackedGameTime = gameTime
    curr.lastAccumulatedDamage = curr.lastAccumulatedDamage + additionalDamage
    If unitKilled Then
        curr.lastUnitsKilled = curr.lastUnitsKilled + 1
        If curr.lastAccumulatedDamage < 1 Then curr.lastAccumulatedDamage = 1 ' std::max<uint64_t>(curr.lastAccumulatedDamage, 1)
    End If
End Sub

' psCommander: a const DROID *
Public Sub addCommanderGroupDamageForCurrentTick(ByVal psCommander As Long, ByVal additionalDamage As Double, ByVal unitKilled As Boolean)
    If aDroid(psCommander).group >= 10 Then
        Exit Sub
    End If
    addGroupDamageForCurrentTick aDroid(psCommander).group, additionalDamage, unitKilled
End Sub

Public Sub selectGroup(ByVal groupNumber As Long)
    ' select the group
    kf_SelectGrouping groupNumber
End Sub

Public Sub assignSelectedDroidsToGroup(ByVal groupNumber As Long)
    assignObjectToGroup selectedPlayer, groupNumber, True
End Sub

Public Sub updateData()
    ' struct AccumulatedGroupInfo, std::array<AccumulatedGroupInfo, 10> calculatedGroupInfo: an array per field
    Dim numberInGroup(0 To 9) As Long, numberDamagedInGroup(0 To 9) As Long, numberCommandedByGroup(0 To 9) As Long
    Dim totalGroupMaxHealth(0 To 9) As Double, displayDroid(0 To 9) As Long
    Dim unitcounter(0 To 9) As Object  ' std::map<std::vector<uint32_t>, size_t>: a Scripting.Dictionary keyed by the components joined
    Dim most_droids_of_same_type_in_group(0 To 9) As Long
    Dim it As Long, psDroid As Long, groupIdx As Long, idx As Long, psCommander As Long
    Dim components() As Long, n As Long, k As Long, key As String, storedGroupInfo As GroupDisplayInfo

    For idx = 0 To 9
        Set unitcounter(idx) = CreateObject("Scripting.Dictionary")
    Next

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).repairGroup < 10 Then
            numberDamagedInGroup(aDroid(psDroid).repairGroup) = numberDamagedInGroup(aDroid(psDroid).repairGroup) + 1
        End If

        groupIdx = aDroid(psDroid).group
        If aDroid(psDroid).group >= 10 Then
            psCommander = 0
            If aDroid(psDroid).psGroup <> 0 Then psCommander = aGroup(aDroid(psDroid).psGroup).psCommander
            If psCommander <> 0 Then
                If aDroid(psCommander).group < 10 Then
                    groupIdx = aDroid(psCommander).group ' accumulate in commander's group
                    numberCommandedByGroup(groupIdx) = numberCommandedByGroup(groupIdx) + 1
                Else
                    GoTo continue_droid
                End If
            Else
                GoTo continue_droid
            End If
        End If

        ' auto& currGroupInfo = calculatedGroupInfo[groupIdx];

        ' display whatever unit occurs the most in this group
        ' find the identifier for this droid
        n = buildComponentsFromDroid(psDroid, components)
        key = ""
        For k = 0 To n - 1
            key = key & CStr(components(k)) & ","
        Next
        unitcounter(groupIdx).Item(key) = unitcounter(groupIdx).Item(key) + 1
        If unitcounter(groupIdx).Item(key) > most_droids_of_same_type_in_group(groupIdx) Then
            most_droids_of_same_type_in_group(groupIdx) = unitcounter(groupIdx).Item(key)
            displayDroid(groupIdx) = psDroid
        End If
        numberInGroup(groupIdx) = numberInGroup(groupIdx) + 1
        totalGroupMaxHealth(groupIdx) = totalGroupMaxHealth(groupIdx) + aDroid(psDroid).originalBody
continue_droid:
        it = lnNext(it)
    Loop

    For idx = 0 To 9
        Set storedGroupInfo = groupInfo(idx)
        storedGroupInfo.numberInGroup = numberInGroup(idx)
        storedGroupInfo.numberDamagedInGroup = numberDamagedInGroup(idx)
        storedGroupInfo.numberCommandedByGroup = numberCommandedByGroup(idx)
        storedGroupInfo.totalGroupMaxHealth = totalGroupMaxHealth(idx)
        If numberInGroup(idx) > 0 Then
            ' generate a DROID_TEMPLATE from the displayDroid
            templateSetParts displayDroid(idx), apsTemplates(storedGroupInfo.displayDroidTemplate)
        End If
    Next
End Sub
