VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OriginalCampaignBalanceChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/campaign.cpp: CampaignStartOptionsForm's struct OriginalCampaignBalanceChoice
Option Explicit

Public displayName As String               ' WzString
Public description As String               ' WzString
Public modInfo As WzCampaignModInfo        ' optional<WzCampaignModInfo> (Nothing: nullopt)
