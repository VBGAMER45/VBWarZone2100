VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGuideLinkedTopicButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzGuideLinkedTopicButton
' src/screens/guidescreen.cpp: WzGuideLinkedTopicButton's fields (guidescreen_c); the initialisers are WzGuideLinkedTopicButton_construct's
' and Class_Initialize's.
Option Explicit

Public font As Long                          ' iV_fonts = font_regular
Public displayName As String                 ' WzString
Public cachedText As WzCachedText
Public horizontalPadding As Long             ' int32_t = 3
Public leftMargin As Long                    ' int32_t = 0
Public ptsAboveBase As Long                  ' int32_t = 0

Private Sub Class_Initialize()
    Set cachedText = New WzCachedText
End Sub
