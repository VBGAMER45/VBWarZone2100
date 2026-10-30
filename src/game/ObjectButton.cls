VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class ObjectButton
' src/hci/objects_stats.h: ObjectButton's fields (objects_stats_c)
Option Explicit

Public objectIndex As Long              ' size_t

Private jumpPosition_ As Vector2i       ' = {0, 0}

Friend Property Get jumpPosition() As Vector2i
    jumpPosition = jumpPosition_
End Property

Friend Property Let jumpPosition(v As Vector2i)
    jumpPosition_ = v
End Property
