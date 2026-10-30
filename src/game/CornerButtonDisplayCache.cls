VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CornerButtonDisplayCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/widget/button.cpp: struct CornerButtonDisplayCache (button_c's makeFormTransparentCornerButton
' sets buttonBackgroundFill: CornerButtonDisplayCache(PIELIGHT buttonBackgroundFill)).
Option Explicit

Public buttonBackgroundFill As Long     ' PIELIGHT
Public text As WzText

Private Sub Class_Initialize()
    Set text = New WzText
End Sub
