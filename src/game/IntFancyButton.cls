VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntFancyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class IntFancyButton
' src/intdisplay.h: IntFancyButton's fields (intdisplay_c); the anonymous struct model is the
' model_ fields, its Vector3i members through Friend properties
Option Explicit

Private modelPosition_ As Vector3i           ' model.position
Private modelRotation_ As Vector3i           ' model.rotation
Public model_scale As Long              ' model.scale
Public model_rate As Long               ' model.rate
Public buttonType As IntFancyButton_ButtonType ' TOPBUTTON is square, BTMBUTTON has a little up arrow.
Public buttonBackgroundEmpty As Boolean ' = false

Friend Property Get model_position() As Vector3i
    model_position = modelPosition_
End Property

Friend Property Let model_position(v As Vector3i)
    modelPosition_ = v
End Property

Friend Property Get model_rotation() As Vector3i
    model_rotation = modelRotation_
End Property

Friend Property Let model_rotation(v As Vector3i)
    modelRotation_ = v
End Property
