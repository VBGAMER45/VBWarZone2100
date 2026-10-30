VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PowerBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class PowerBar
' src/intdisplay.h: PowerBar's fields (intdisplay_c): struct DisplayPowerBarCache cache (wzText is
' wzText_: VB would take it for the class)
Option Explicit

Public wzText_ As WzText                ' cache.wzText
Public wzNeedText As WzText             ' cache.wzNeedText

Private Sub Class_Initialize()
    Set wzText_ = New WzText
    Set wzNeedText = New WzText
End Sub
