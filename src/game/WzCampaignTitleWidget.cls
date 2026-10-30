VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCampaignTitleWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzCampaignTitleWidget
' src/titleui/campaign.cpp: the fields of class WzCampaignTitleWidget (campaign_title_c); the initialisers are its
' _construct's.
Option Explicit

Public leftPadding As Long                 ' = 40
Public topPadding As Long                  ' = 20
Public titleAuthorPadding As Long          ' = 3
Public pBackgroundTexture As gfx_api_texture   ' gfx_api::texture* = nullptr
Public nameLabel As WIDGET                 ' std::shared_ptr<W_LABEL>
