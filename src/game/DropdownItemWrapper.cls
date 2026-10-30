VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DropdownItemWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class DropdownItemWrapper
' lib/widget/dropdown.h: DropdownItemWrapper's fields (dropdown_c). std::weak_ptr<DropdownWidget> parent
' is the dropdown's serial (widgweak_vb).
Option Explicit

Public weakSerial As Long
Public item As WIDGET
Public onSelect As WzFunc               ' DropdownOnSelectHandler
Public parent As Long
Public selected As Boolean              ' = false
Public mouseDownOnWrapper As Boolean    ' = false
