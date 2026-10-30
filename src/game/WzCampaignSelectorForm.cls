VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCampaignSelectorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzCampaignSelectorForm
' src/titleui/campaign.cpp: the fields of class WzCampaignSelectorForm (campaign_title_c); the initialisers are its
' _construct's.
Option Explicit

Public parentCampaignSelectorUI As Long    ' std::weak_ptr<WzCampaignSelectorTitleUI>: its address
Public originalCampaignButton As WIDGET    ' std::shared_ptr<WzCampaignModButton>
Public alternateCampaigns As WIDGET        ' std::shared_ptr<WzCampaignListSection>
