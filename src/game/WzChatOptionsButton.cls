VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzChatOptionsButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzChatOptionsButton
' src/hci/chatoptions.cpp: the fields of class WzChatOptionsButton (chatoptions_c); the initialisers are its _construct's.
Option Explicit

Public wzMessageText As WzText
Public lastWidgetWidth As Long             ' = 0
Public isTruncated As Boolean              ' = false

Private Sub Class_Initialize()
    Set wzMessageText = New WzText
End Sub
