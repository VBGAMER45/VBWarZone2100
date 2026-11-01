VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzBlindPlayersReadyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzBlindPlayersReadyInfo
' src/titleui/widgets/blindwaitingroom.cpp: the fields of class WzBlindPlayersReadyInfo (blindwaitingroom_c); the initialisers are its _construct's.
Option Explicit

Public topPadding As Long               ' = 4
Public bottomPadding As Long            ' = 4
Public internalHorizontalPadding As Long ' = 10
Public readyCountLabel As WIDGET        ' std::shared_ptr<W_LABEL>
Public readyCountValue As WIDGET        ' std::shared_ptr<W_LABEL>
Public lastStatusUpdate As Double       ' std::chrono::steady_clock::time_point (widgSteadyNowMs; 0: the epoch)
Public minUpdateFreq As Double          ' std::chrono::milliseconds = 500
Public tooltip As String                ' std::string
