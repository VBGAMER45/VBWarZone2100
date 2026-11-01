VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzBlindRoomPlayerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzBlindRoomPlayerInfo
' src/titleui/widgets/blindwaitingroom.cpp: the fields of class WzBlindRoomPlayerInfo (blindwaitingroom_c); the initialisers are its _construct's.
Option Explicit

Public topPadding As Long               ' = 4
Public bottomPadding As Long            ' = 4
Public internalVerticalPadding As Long  ' = 10
Public playerLabel As WIDGET            ' std::shared_ptr<W_LABEL>
Public playerRow As WIDGET              ' std::shared_ptr<WzPlayerRow>
