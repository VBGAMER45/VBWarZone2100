VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCampaignModButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzCampaignModButton
' src/titleui/campaign.cpp: the fields of class WzCampaignModButton (campaign_title_c); the initialisers are its
' _construct's.
Option Explicit

Public wzText As WzText
Public isSelected As Boolean                ' = false

Private Sub Class_Initialize()
    Set wzText = New WzText
End Sub
