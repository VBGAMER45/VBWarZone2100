VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMultiButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/multiint.h: the fields of class WzMultiButton (multiint_lobby_c)
' (VB) data of widget class WzMultiButton
Option Explicit

Private imNormal_ As AtlasImage
Private imDown_ As AtlasImage
Public doHighlight As Long              ' unsigned
Public tc As Long                       ' unsigned
Public alpha As Long                    ' uint8_t = 255
Public downStateMask As Long            ' unsigned = WBUT_DOWN | WBUT_LOCK | WBUT_CLICKLOCK
Public greyStateMask As Long            ' unsigned = WBUT_DISABLE
Public has_drawBlueBorder As Boolean    ' optional<bool> drawBlueBorder: has_value()
Public drawBlueBorder As Boolean        ' its value

Friend Property Get imNormal() As AtlasImage
    imNormal = imNormal_
End Property

Friend Property Let imNormal(v As AtlasImage)
    imNormal_ = v
End Property

Friend Property Get imDown() As AtlasImage
    imDown = imDown_
End Property

Friend Property Let imDown(v As AtlasImage)
    imDown_ = v
End Property
