VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TeamStrategyView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class TeamStrategyView
' src/hci/teamstrategy.cpp: the fields of class TeamStrategyView (teamstrategy_c); the initialisers are its _construct's.
Option Explicit

Public description As WIDGET            ' std::shared_ptr<W_LABEL>
Public instructions As WIDGET           ' std::shared_ptr<W_LABEL>
Public table As WIDGET                  ' std::shared_ptr<ScrollableTableWidget>
Public columnImagesManager As WzTeamStratImagesManager ' std::shared_ptr<WzTeamStrategyColumnImagesManager>
Public noTeammatesLabel As WIDGET       ' std::shared_ptr<W_LABEL>
Public playerNameColIdx As Long         ' size_t = 0
Public updatingPlayers As Boolean       ' = false
Public backgroundColor As Long          ' PIELIGHT = pal_RGBA(0,0,0,0)
