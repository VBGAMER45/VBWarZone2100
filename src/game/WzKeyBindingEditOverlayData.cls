VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzKeyBindingEditOverlayData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/options/controlsoptions.cpp: the fields of class WzFullscreenKeyBindingEditingOverlay
' (VB) data of widget class WzFullscreenKeyBindingEditingOverlay
Option Explicit

Public onPushedKeyCombo As WzFunc        ' OnPushedKeyComboFunc: void(const KeyMappingInput input, optional<KEY_CODE> metaKey)
Public metakeyDown_RALT As Boolean       ' = false
Public metakeyDown_LALT As Boolean
Public metakeyDown_RCTRL As Boolean
Public metakeyDown_LCTRL As Boolean
Public metakeyDown_RSHIFT As Boolean
Public metakeyDown_LSHIFT As Boolean
Public metakeyDown_RMETA As Boolean
Public metakeyDown_LMETA As Boolean
Public mouseDownOutsideTargetWidget As Boolean ' = false
