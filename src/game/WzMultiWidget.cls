VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMultiWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/multimenu.cpp: the fields of class WzMultiWidget (multimenu_c); panels is a Collection (element n is
' item n + 1); weakSerial: (VB) its std::weak_ptrs (the tab switcher's task).
' (VB) data of widget class WzMultiWidget
Option Explicit

Public weakSerial As Long
Public closeButton As WIDGET            ' std::shared_ptr<W_BUTTON>
Public panelSwitcher As WIDGET          ' std::shared_ptr<WzMultiMenuTabs>
Public multiMenuGrid As WIDGET          ' std::shared_ptr<MultiMenuGrid>
Public teamStrategyViewWrapped As WIDGET
Public panels As Collection             ' std::vector<std::shared_ptr<WIDGET>>
Public currentlyAttachedPanel As WIDGET

Private Sub Class_Initialize()
    Set panels = New Collection
End Sub
