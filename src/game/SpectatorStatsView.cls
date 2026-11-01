VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpectatorStatsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class SpectatorStatsView
' src/spectatorwidgets.cpp: the fields of class SpectatorStatsView (spectatorwidgets_c); the initialisers are its _construct's.
Option Explicit

Public optionsButton As WIDGET          ' std::shared_ptr<W_BUTTON> (never set)
Public closeButton As WIDGET            ' std::shared_ptr<W_BUTTON>
Public table As WIDGET                  ' std::shared_ptr<ScrollableTableWidget>
Public weaponGradesManager As WIDGET    ' std::shared_ptr<WzWeaponGradesColumnManager>
Public playerNameColIdx As Long         ' size_t = 1
