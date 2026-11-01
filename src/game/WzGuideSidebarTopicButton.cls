VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGuideSidebarTopicButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzGuideSidebarTopicButton
' src/screens/guidescreen.cpp: WzGuideSidebarTopicButton's fields (guidescreen_c); the initialisers are WzGuideSidebarTopicButton_construct's
' and Class_Initialize's.
Option Explicit

Public topicIdentifier As String             ' std::string
Public indentLevel As Long                   ' size_t = 0
Public wzText As WzCachedText
Public isSelected As Boolean                 ' = false
Public isNew As Boolean                      ' = true
Public isTruncated As Boolean                ' = false

Private Sub Class_Initialize()
    Set wzText = New WzCachedText
    isNew = True
End Sub
