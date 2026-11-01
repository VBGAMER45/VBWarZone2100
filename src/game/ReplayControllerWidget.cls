VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReplayControllerWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class ReplayControllerWidget
' src/hci.cpp: the fields of class ReplayControllerWidget (hci_replay_c); the initialisers are its _construct's.
Option Explicit

Public titleLabel As WIDGET             ' std::shared_ptr<W_LABEL>
Public pauseButton As WIDGET            ' std::shared_ptr<W_REPLAY_CONTROL_BUTTON>
Public playButton As WIDGET             ' std::shared_ptr<W_REPLAY_CONTROL_BUTTON>
Public fastForwardButton As WIDGET      ' std::shared_ptr<W_REPLAY_CONTROL_BUTTON>
Public lastUpdateTime As Long           ' UDWORD = 0
Public autoStateRefreshInterval As Long ' UDWORD = GAME_TICKS_PER_SEC * 3
