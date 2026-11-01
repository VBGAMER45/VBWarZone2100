VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzBlindRoomStatusInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzBlindRoomStatusInfo
' src/titleui/widgets/blindwaitingroom.cpp: the fields of class WzBlindRoomStatusInfo (blindwaitingroom_c); the initialisers are its _construct's.
Option Explicit

Public topPadding As Long               ' = 4
Public bottomPadding As Long            ' = 4
Public internalVerticalPadding As Long  ' = 5
Public statusLabel As WIDGET            ' std::shared_ptr<W_LABEL>
Public statusDetails As WIDGET          ' std::shared_ptr<W_LABEL>
Public indeterminateIndicator As WIDGET ' std::shared_ptr<WIDGET>
Public statusStr As String              ' WzString
Public lastStatusUpdate As Double       ' std::chrono::steady_clock::time_point (widgSteadyNowMs; 0: the epoch)
Public minUpdateFreq As Double          ' std::chrono::milliseconds = 500
