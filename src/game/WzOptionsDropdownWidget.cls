VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzOptionsDropdownWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/optionsform.h: the fields of class WzOptionsDropdownWidget
' (VB) data of widget class WzOptionsDropdownWidget
Option Explicit

Public currentOptionChoiceDisplayWidg As WIDGET ' std::shared_ptr<WzOptionsChoiceWidget>
Public paddingLeft As Long               ' = 4
Public isHighlighted As Boolean          ' = false
