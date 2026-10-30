VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GroupDisplayInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/hci/groups.cpp: struct GroupsUIController::GroupDisplayInfo. displayDroidTemplate (a
' DROID_TEMPLATE member) is a template of template_c's pool, the object's own (its handle).
' A uint64_t is a Double.
Option Explicit

Public numberInGroup As Long            ' size_t = 0
Public numberDamagedInGroup As Long     ' size_t = 0
Public numberCommandedByGroup As Long   ' size_t = 0: the number of droids commanded by commanders in this group
Public totalGroupMaxHealth As Double    ' uint64_t = 0
Public displayDroidTemplate As Long     ' DROID_TEMPLATE (the pool handle)

Public currAttackGlowAlpha As Long      ' uint8_t = 0

' used for calculating display info
Public lastAttackedGameTime As Long     ' uint32_t = 0
Public lastAccumulatedDamage As Double  ' uint64_t = 0
Public lastUnitsKilled As Double        ' uint64_t = 0

Private Sub Class_Initialize()
    displayDroidTemplate = newTemplate()
End Sub

Private Sub Class_Terminate()
    deleteTemplate displayDroidTemplate
End Sub
