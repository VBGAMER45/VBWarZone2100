VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_REPLAY_CONTROL_BUTTON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_REPLAY_CONTROL_BUTTON
' src/hci.cpp: the fields of class W_REPLAY_CONTROL_BUTTON (hci_replay_c); the initialisers are its _construct's.
Option Explicit

Public pButtonImage As gfx_api_texture  ' gfx_api::texture* = nullptr (deleted with the button: released with this object)
Public outline As Long                  ' OutlineState = OutlineState::None
