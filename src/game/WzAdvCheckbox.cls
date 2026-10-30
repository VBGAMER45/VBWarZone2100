VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzAdvCheckbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzAdvCheckbox
' src/titleui/widgets/advcheckbox.h: WzAdvCheckbox's fields (advcheckbox_c); the initialisers are
' WzAdvCheckbox_construct's.
Option Explicit

Public descriptionWidget As WIDGET      ' std::shared_ptr<Paragraph>
Public wzText As WzText
Public imageDimensions As Long
Public cachedIdealWidth As Long
Public lastWidgetWidth As Long
Public outerHorizontalPadding As Long
Public innerHorizontalPadding As Long
Public outerVerticalPadding As Long
Public innerVerticalPadding As Long
Public checked As Boolean

Private Sub Class_Initialize()
    Set wzText = New WzText
End Sub
