VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzOptionsChoiceWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/optionsform.h: the fields of class WzOptionsChoiceWidget
' (VB) data of widget class WzOptionsChoiceWidget
Option Explicit

Public text As String
Public wzText As WzCachedText
Public FontID As iV_fonts
Public icon As Long                      ' const AtlasImageDef *: a handle (0 = nullptr)
Public iconDrawWidth As Long             ' = 0
Public iconDrawHeight As Long            ' = 0
Public highlightBackgroundColor As Long  ' PIELIGHT
Public lastWidgetWidth As Long           ' = 0
Public horizontalPadding As Long         ' = 10
Public iconPaddingAfter As Long          ' = 5
Public verticalPadding As Long           ' = 5
Public isHighlighted As Boolean          ' = false
Public forceDisplayHighlight As Boolean  ' = false
Public isDisabled As Boolean             ' = false
Public isTruncated As Boolean            ' = false
