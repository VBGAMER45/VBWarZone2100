VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WZOptionTitleLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WZOptionTitleLabel
' src/titleui/campaign.cpp: the fields of class WZOptionTitleLabel (campaign_title_c); the initialisers are its
' _construct's.
Option Explicit

Public wzText As WzText
Public FontID As iV_fonts                  ' = font_regular
Public verticalPadding As Long             ' = 5
Public isTruncated As Boolean              ' = false

Private Sub Class_Initialize()
    Set wzText = New WzText
End Sub
