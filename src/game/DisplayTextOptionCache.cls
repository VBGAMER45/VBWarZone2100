VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisplayTextOptionCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/frontend.h: struct DisplayTextOptionCache, the pUserData of the front end's text widgets
' (displayTextOption, displayTextAt270): their WzText and how it shrinks to fit. A new one is
' C++'s new DisplayTextOptionCache(); the widget's onDelete drops it (pUserData = Empty).
Option Explicit

Public overflowBehavior_ As OverflowBehavior   ' = OverflowBehavior::None
Public lastWidgetWidth As Long                ' = 0
Public wzText As WzText

Private Sub Class_Initialize()
    Set wzText = New WzText
End Sub
