VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzPlayerBoxTabButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/lobbyplayerboxtabs.cpp: the fields of class WzPlayerBoxTabButton
' (VB) data of widget class WzPlayerBoxTabButton
Option Explicit

Public numTakenSlots As Long            ' uint8_t
Public numTakenSlotsReady As Long       ' uint8_t
Public totalAvailableSlots As Long      ' uint8_t
Public rightPadding As Long
Public selected As Boolean
Public titleLabel As WIDGET             ' std::shared_ptr<W_LABEL>
Public slotsCountsLabel As WIDGET       ' std::shared_ptr<W_LABEL>
