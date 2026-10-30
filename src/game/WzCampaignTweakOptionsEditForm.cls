VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCampaignTweakOptionsEditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzCampaignTweakOptionsEditForm
' src/titleui/campaign.cpp: the fields of class WzCampaignTweakOptionsEditForm (campaign_title_c); the initialisers are its
' _construct's.
Option Explicit

Public formTitle As WIDGET                 ' std::shared_ptr<W_LABEL>
Public closeButton As WIDGET               ' std::shared_ptr<WzFrontendImageButton>
Public resetButton As WIDGET               ' std::shared_ptr<WzFrontendImageButton>
Public tweaksDisplayList As WIDGET         ' std::shared_ptr<ScrollableListWidget>
Public toggleWidgets As Collection         ' std::vector<std::shared_ptr<WzCampaignTweakOptionToggle>>
Public tweakOptionIdxToToggleWidgetIdxMap As Collection  ' std::unordered_map<size_t, size_t>: "i<index>" -> index
Public tweakOptions As Collection          ' std::shared_ptr<std::vector<WzCampaignTweakOptionSetting>>
Public innerPadding As Long                ' = 20
Public betweenButtonPadding As Long        ' = 15

Private Sub Class_Initialize()
    Set toggleWidgets = New Collection
    Set tweakOptionIdxToToggleWidgetIdxMap = New Collection
End Sub
