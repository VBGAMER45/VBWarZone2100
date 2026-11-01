VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LobbyStatusOverlayWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class LobbyStatusOverlayWidget
' src/titleui/widgets/gamebrowserform.cpp: the fields of class LobbyStatusOverlayWidget (gamebrowserform_c)
Option Explicit

Public textLabel As WIDGET              ' std::shared_ptr<W_LABEL>
Public indeterminateIndicator As WIDGET ' std::shared_ptr<WIDGET>
Public resetLobbyAddressButton As WIDGET ' std::shared_ptr<WzFrontendImageButton>

Public onResetLobbyAddressClickFunc As WzFunc ' OnResetLobbyAddressClickFunc

Public weakSerial As Long               ' (VB) its std::weak_ptr<LobbyStatusOverlayWidget>s: the widgweak_vb serial
