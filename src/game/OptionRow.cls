VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/optionsform.cpp: the fields of class OptionRow
' (VB) data of widget class OptionRow
Option Explicit

Public m_optionId As String
Public m_labelWidget As WIDGET           ' std::shared_ptr<WzOptionsLabelWidget>
Public m_wrappedOptionValueChangerWidget As WIDGET ' std::shared_ptr<OptionsValueChangerWidgetWrapper>
Public m_weakParentForm As Long          ' std::weak_ptr<OptionsForm>: its address
Public m_indentLevel As Long             ' int16_t = 0
Public has_flashHighlightStart As Boolean ' optional<uint32_t> flashHighlightStart = nullopt
Public flashHighlightStart As Long
Public showBottomBorder As Boolean       ' = false
Public isHighlighted As Boolean          ' = false
