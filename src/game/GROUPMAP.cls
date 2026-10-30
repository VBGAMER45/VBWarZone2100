VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GROUPMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/qtscript.cpp, qtscript.h: struct scripting_engine::GROUPMAP, a script
' instance's groups (made with New, as new GROUPMAP). An object is its
' handle. ObjectToGroupMap (std::unordered_map<const BASE_OBJECT *,
' groupID>) is a Scripting.Dictionary of handle -> group id;
' std::unordered_map<groupID, GroupSet> m_groups a Dictionary of group id ->
' Dictionary (the set: handle -> True). Both iterate in insertion order, not
' MSVC's hash order: getGroupObjects (enumGroup) and the saved groups list
' objects in the order they joined.
' The protected members scripting_engine uses (getLastNewGroupId,
' saveLoadSetLastNewGroupId) are Friend.
Option Explicit

Private m_map As Object                 ' ObjectToGroupMap
Private m_groups As Object              ' std::unordered_map<groupID, GroupSet>
Private lastNewGroupId As Long

Private Sub Class_Initialize()
    Set m_map = CreateObject("Scripting.Dictionary")
    Set m_groups = CreateObject("Scripting.Dictionary")
    lastNewGroupId = 0
End Sub

Friend Function getLastNewGroupId() As Long
    getLastNewGroupId = lastNewGroupId
End Function

Friend Sub saveLoadSetLastNewGroupId(ByVal value As Long)
    If Not (value >= 0) Then ASSERT_FAILURE "value >= 0", "src/qtscript.cpp:88", "saveLoadSetLastNewGroupId", "Invalid value: %d", value: Exit Sub
    lastNewGroupId = value
End Sub

' m_groups[groupId] (operator[]: an empty set is made when there is none)
Private Function groupSet(ByVal groupId As Long) As Object
    If Not m_groups.Exists(groupId) Then
        m_groups.Add groupId, CreateObject("Scripting.Dictionary")
    End If
    Set groupSet = m_groups.Item(groupId)
End Function

Public Function newGroupID() As Long
    Dim newId As Long, busy As Boolean
    newId = lastNewGroupId + 1
    Do
        busy = False
        If m_groups.Exists(newId) Then
            busy = m_groups.Item(newId).Count > 0
        End If
        If Not busy Then Exit Do
        If newId < &H7FFFFFFF Then
            newId = newId + 1
        Else
            ' loop back around
            ' NOTE: group zero is reserved
            newId = 1
        End If
    Loop
    lastNewGroupId = newId
    newGroupID = newId
End Function

Public Sub insertObjectIntoGroup(ByVal psObj As Long, ByVal groupId As Long)
    Dim s As Object, inserted As Boolean
    If Not m_map.Exists(psObj) Then
        m_map.Add psObj, groupId
        Set s = groupSet(groupId)
        inserted = Not s.Exists(psObj)
        If inserted Then s.Add psObj, True
        If Not (inserted) Then ASSERT_FAILURE "groupSetResult.second", "src/qtscript.cpp:118", "insertObjectIntoGroup", "Object already exists in group!"
    End If
End Sub

' map(): the objects and their groups (a Dictionary of handle -> group id)
Public Function map_() As Object
    Set map_ = m_map
End Function

Public Function groupSize(ByVal groupId As Long) As Long
    If m_groups.Exists(groupId) Then
        groupSize = m_groups.Item(groupId).Count
        Exit Function
    End If
    groupSize = 0
End Function

' optional<groupID>: True (has_value) with the group in groupId
Public Function removeObjectFromGroup(ByVal psObj As Long, groupId As Long) As Boolean
    Dim numItemsErased As Long, s As Object
    If m_map.Exists(psObj) Then
        groupId = m_map.Item(psObj)
        m_map.Remove psObj

        Set s = groupSet(groupId)
        If s.Exists(psObj) Then
            s.Remove psObj
            numItemsErased = 1
        End If
        If Not (numItemsErased = 1) Then ASSERT_FAILURE "numItemsErased == 1", "src/qtscript.cpp:141", "removeObjectFromGroup", "Object did not exist in group set??"
        removeObjectFromGroup = True
        Exit Function
    End If
    removeObjectFromGroup = False
End Function

' std::vector<const BASE_OBJECT *>: the count, the objects in result(0 .. count - 1)
Public Function getGroupObjects(ByVal groupId As Long, result() As Long) As Long
    Dim s As Object, k As Variant, n As Long
    If m_groups.Exists(groupId) Then
        Set s = m_groups.Item(groupId)
        If s.Count > 0 Then
            ReDim result(0 To s.Count - 1)
            For Each k In s.Keys
                result(n) = k
                n = n + 1
            Next
        End If
    End If
    getGroupObjects = n
End Function
