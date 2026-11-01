VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzLoadingPlayersStatusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzLoadingPlayersStatusForm
' src/screens/netpregamescreen.cpp: the fields of class WzLoadingPlayersStatusForm (netpregamescreen_c); the initialisers are its _construct's.
Option Explicit

Public titleLabel As WIDGET             ' std::shared_ptr<W_LABEL>
Public playersList As WIDGET            ' std::shared_ptr<ScrollableListWidget>
Public spectatorsList As WIDGET         ' std::shared_ptr<ScrollableListWidget>
Public lastUpdateTime As Long           ' UDWORD = 0
Public autoStateRefreshInterval As Long ' UDWORD = GAME_TICKS_PER_SEC * 3

Private Sub Class_Initialize()
    autoStateRefreshInterval = GAME_TICKS_PER_SEC * 3
End Sub
