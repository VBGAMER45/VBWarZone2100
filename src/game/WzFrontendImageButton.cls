VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzFrontendImageButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzFrontendImageButton
' src/titleui/widgets/frontendimagebutton.h: WzFrontendImageButton's fields (frontendimagebutton_c);
' optional<UWORD> / optional<PIELIGHT> are a has_ flag and the value, the initialisers are
' WzFrontendImageButton_construct's.
Option Explicit

Public wzText As WzText
Public has_frontendImgID As Boolean
Public frontendImgID As Long            ' UWORD
Public has_customTextColor As Boolean
Public customTextColor As Long          ' PIELIGHT
Public has_customTextHighlightColor As Boolean
Public customTextHighlightColor As Long
Public has_customImgColor As Boolean
Public customImgColor As Long
Public has_customBackgroundColor As Boolean
Public customBackgroundColor As Long
Public missingImage As Boolean
Public imageDimensions As Long
Public cachedIdealWidth As Long
Public lastWidgetWidth As Long
Public horizontalPadding As Long
Public verticalPadding As Long
Public imageHorizontalOffset As Long
Public borderDrawMode_ As BorderDrawMode

Private Sub Class_Initialize()
    Set wzText = New WzText
End Sub
