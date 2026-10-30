VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionsKeyBindingWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/options/controlsoptions.cpp: the fields of class OptionsKeyBindingWidget
' (VB) data of widget class OptionsKeyBindingWidget
Option Explicit

Public text As String
Public wzText As WzCachedText
Public FontID As iV_fonts                ' = font_regular (hides W_BUTTON::FontID)
Public bindingColor As Long              ' PIELIGHT = WZCOL_FORM_TEXT
Public cachedIdealTextWidth As Long      ' = 0
Public bEditable As Boolean              ' = false
Public isSelected As Boolean             ' = false
Public isEditingMode As Boolean          ' = false
Public isTruncated As Boolean            ' = false
Public lastWidgetWidth As Long           ' = 0
Public highlightBackgroundColor As Long  ' PIELIGHT
