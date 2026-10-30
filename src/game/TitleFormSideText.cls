VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TitleFormSideText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class TitleFormSideText
' src/titleui/widgets/titleformwrapper.cpp: TitleFormSideText's fields (titleformwrapper_c); the
' initialisers are TitleFormSideText_construct's.
Option Explicit

Public text As String
Public fontID As Long                   ' iV_fonts
Public wzText As WzText
Public horizontalPadding As Long
Public verticalPadding As Long
Public lastWidgetWidth As Long
Public isTruncated As Boolean

Private Sub Class_Initialize()
    Set wzText = New WzText
End Sub
