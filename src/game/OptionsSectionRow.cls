VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionsSectionRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/optionsform.cpp: the fields of class OptionsSectionRow
' (VB) data of widget class OptionsSectionRow
Option Explicit

Public m_labelWidget As WIDGET           ' std::shared_ptr<W_LABEL>
Public showBottomBorder As Boolean       ' = false
Public has_flashHighlightStart As Boolean ' optional<uint32_t> flashHighlightStart = nullopt
Public flashHighlightStart As Long
