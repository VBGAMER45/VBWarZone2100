VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzJoiningIndicatorData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzJoiningIndeterminateIndicatorWidget
' src/screens/joiningscreen.cpp: the fields of class WzJoiningIndeterminateIndicatorWidget (joiningscreen_c; this
' class's name: VB's 31 characters)
Option Explicit

Public periodText As WzText             ' WzText periodText;

Private Sub Class_Initialize()
    Set periodText = New WzText
End Sub
