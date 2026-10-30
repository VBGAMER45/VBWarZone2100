VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionsBrowserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/options/optionsforms.cpp: the fields of class OptionsBrowserForm
' (VB) data of widget class OptionsBrowserForm
Option Explicit

Public weakSerial As Long               ' (VB) its std::weak_ptr serial (widgweak_vb)
Private optionsForms_() As OptionsBrowserForm_OptionsFormDetails ' std::vector<OptionsFormDetails>
Public optionsFormsCount As Long
Public has_currentOptionsForm As Boolean ' optional<size_t> currentOptionsForm = nullopt
Public currentOptionsForm As Long
Public optionalBackButton As WIDGET
Public optionsFormSwitcher As WIDGET
Public optionsFormSectionSwitcher As WIDGET
Public openConfigurationDirectoryLink As WIDGET
Public maxOptionsFormWidth As Long       ' = 450

Friend Property Get optionsForms(ByVal i As Long) As OptionsBrowserForm_OptionsFormDetails
    optionsForms = optionsForms_(i)
End Property

Friend Property Let optionsForms(ByVal i As Long, v As OptionsBrowserForm_OptionsFormDetails)
    optionsForms_(i) = v
End Property

Friend Sub optionsForms_push_back(v As OptionsBrowserForm_OptionsFormDetails)
    If optionsFormsCount = 0 Then
        ReDim optionsForms_(0 To 7)
    ElseIf optionsFormsCount > UBound(optionsForms_) Then
        ReDim Preserve optionsForms_(0 To optionsFormsCount * 2 - 1)
    End If
    optionsForms_(optionsFormsCount) = v
    optionsFormsCount = optionsFormsCount + 1
End Sub
