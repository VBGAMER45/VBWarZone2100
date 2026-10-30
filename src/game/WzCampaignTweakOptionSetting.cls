VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCampaignTweakOptionSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/campaign.cpp: class WzCampaignTweakOptionSetting (its constructor: campaign_title_c's
' WzCampaignTweakOptionSetting_new)
Option Explicit

Public uniqueIdentifier As String
Public displayName As String               ' WzString
Public description As String               ' WzString
Public defaultValue As Boolean             ' = false
Public currentValue As Boolean             ' = false
Public editable As Boolean                 ' = true

Private Sub Class_Initialize()
    editable = True
End Sub
