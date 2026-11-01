VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGuideTopicsNavButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzGuideTopicsNavButton
' src/screens/guidescreen.cpp: WzGuideTopicsNavButton's fields (guidescreen_c); the initialisers are WzGuideTopicsNavButton_construct's
' and Class_Initialize's.
Option Explicit

Public wzMessageText As WzText
Public has_intfacImageID As Boolean          ' optional<UWORD> intfacImageID = nullopt
Public intfacImageID As Long
Public imageDimensions As Long               ' = 16
Public lastWidgetWidth As Long               ' = 0
Public isTruncated As Boolean                ' = false

Private Sub Class_Initialize()
    Set wzMessageText = New WzText
    imageDimensions = 16
End Sub
