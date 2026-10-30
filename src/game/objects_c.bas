Attribute VB_Name = "objects_c"
' src/objects.cpp, objects.h: the object system.
Option Explicit

' Initialise the object system
Public Function objInitialise() As Boolean
    If Not objmemInitialise() Then
        Exit Function
    End If

    objInitialise = True
End Function

' Shutdown the object system
Public Function objShutdown() As Boolean
    objmemShutdown

    objShutdown = True
End Function

Public Function objInfo(ByVal psObj As Long) As String
    If psObj = 0 Then
        objInfo = "null"
        Exit Function
    End If

    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        objInfo = droidGetName(psObj)
    Case OBJ_STRUCTURE
        objInfo = asStructureStats(aStruct(psObj).pStructureType).name_
    Case OBJ_FEATURE
        objInfo = asFeatureStats(aFeature(psObj).psStats).name_
    Case OBJ_PROJECTILE
        objInfo = "Projectile"  ' TODO
    Case Else
        objInfo = "Unknown object type"
    End Select
End Function
