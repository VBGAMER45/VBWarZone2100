VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionsButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/options/audiooptions.cpp: the fields of class OptionsButton
' (VB) data of widget class OptionsButton
Option Explicit

Public wzText As WzCachedText
Public FontID As iV_fonts                ' = font_regular (hides W_BUTTON::FontID)
Public bindingColor As Long              ' PIELIGHT = WZCOL_FORM_TEXT
Public highlightBackgroundColor As Long  ' PIELIGHT
Public cachedIdealTextWidth As Long      ' = 0
Public lastWidgetWidth As Long           ' = 0
Public verticalPadding As Long           ' = 10
Public horizontalPadding As Long         ' = 14
Public isTruncated As Boolean            ' = false
Public updateButtonFunc As WzFunc        ' UpdateButtonFunc: void(OptionsButton&)
Public needsStateUpdates As Boolean
