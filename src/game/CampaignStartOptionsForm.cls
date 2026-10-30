VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CampaignStartOptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class CampaignStartOptionsForm
' src/titleui/campaign.cpp: the fields of class CampaignStartOptionsForm (campaign_title_c); the initialisers are its
' _construct's.
Option Explicit

Public overlayScreen As W_SCREEN           ' std::shared_ptr<W_SCREEN>
Public overlayScreenOnCloseFunc As WzFunc  ' std::function<void ()>
Public titleArea As WIDGET                 ' std::shared_ptr<WzCampaignTitleWidget>
Public subInfoBox As WIDGET                ' std::shared_ptr<WzCampaignSubInfoBox>
Public modInfo As WzCampaignModInfo        ' optional<WzCampaignModInfo> (Nothing: nullopt)
Public campaignBalanceDropdown As WIDGET   ' std::shared_ptr<DropdownWidget>
Public campaignSelectorDropdown As WIDGET  ' std::shared_ptr<DropdownWidget>
Public tweakOptionsSummary As WIDGET       ' std::shared_ptr<WzCampaignTweakOptionsSummaryWidget>
' std::vector<CAMPAIGN_FILE> campaignList: campaignList(0 .. campaignList_size - 1)
Private campaignList_() As CAMPAIGN_FILE
Public campaignList_size As Long
Public campaignBalanceChoices As Collection    ' std::vector<OriginalCampaignBalanceChoice>
Public tweakOptionSettings As Collection   ' std::shared_ptr<std::vector<WzCampaignTweakOptionSetting>>

Private Sub Class_Initialize()
    Set campaignBalanceChoices = New Collection
    Set tweakOptionSettings = New Collection
End Sub

Friend Property Get campaignList(ByVal i As Long) As CAMPAIGN_FILE
    campaignList = campaignList_(i)
End Property

' campaignList = list(0 .. n - 1)
Friend Sub setCampaignList(list() As CAMPAIGN_FILE, ByVal n As Long)
    Dim i As Long
    campaignList_size = n
    If n > 0 Then
        ReDim campaignList_(0 To n - 1)
        For i = 0 To n - 1
            campaignList_(i) = list(i)
        Next
    Else
        Erase campaignList_
    End If
End Sub
