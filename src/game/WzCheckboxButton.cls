VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCheckboxButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzCheckboxButton
' lib/widget/checkbox.h: WzCheckboxButton's fields (checkbox_c); the initialisers are
' WzCheckboxButton_construct's.
Option Explicit

' private
Public wzText As WzText
Public isChecked As Boolean             ' = false
Public textColor As Long                ' PIELIGHT = WZCOL_TEXT_MEDIUM

Private Sub Class_Initialize()
    Set wzText = New WzText
End Sub
