VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiMenuGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/multimenu.cpp: the fields of class MultiMenuGrid (multimenu_c). playersWidgets is a Collection of
' struct PlayerWidgets, each a Variant array (player, name, score, kills, units, lastColumn: the player's
' number and its five W_LABELs); the initialisers are its _construct's.
' (VB) data of widget class MultiMenuGrid
Option Explicit

Public timerLabel As WIDGET             ' std::shared_ptr<W_LABEL>
Public powerLabel As WIDGET
Public pingLabel As WIDGET
Public structsLabel As WIDGET
Public playersWidgets As Collection     ' std::vector<PlayerWidgets>

Private Sub Class_Initialize()
    Set playersWidgets = New Collection
End Sub
