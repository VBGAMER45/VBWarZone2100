VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntFormAnimated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class IntFormAnimated
' src/intdisplay.h: IntFormAnimated's fields (intdisplay_c)
Option Explicit

Public startTime As Long                ' unsigned: Animation start time
Public currentAction As Long            ' Opening/open/closing/closed.
Public onCloseAnimFinished As WzFunc    ' W_ANIMATED_ON_CLOSE_FUNC (Nothing = nullptr)
