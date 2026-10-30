VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCampaignTweakOptSummaryData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzCampaignTweakOptionsSummaryWidget
' src/titleui/campaign.cpp: the fields of class WzCampaignTweakOptionsSummaryWidget (campaign_title_c); the initialisers are its
' _construct's.
' (VB) named WzCampaignTweakOptSummaryData: the class name has over 31 characters
Option Explicit

Public tweakOptionsList As WIDGET          ' std::shared_ptr<ClickableScrollableList>
Public tweakOptionsLabels As Collection    ' std::vector<std::shared_ptr<W_LABEL>>
Public infoButton As WIDGET                ' std::shared_ptr<WzInfoButton>
Public infoClickHandler As WzFunc          ' InfoClickHandler (Nothing: empty)
Public betweenRowPadding As Long           ' = 0
Public topPadding As Long                  ' = 8
Public bottomPadding As Long               ' = 5
Public internalHorizontalPadding As Long   ' = 10

Private Sub Class_Initialize()
    Set tweakOptionsLabels = New Collection
End Sub
