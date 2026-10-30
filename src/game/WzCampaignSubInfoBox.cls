VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCampaignSubInfoBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzCampaignSubInfoBox
' src/titleui/campaign.cpp: the fields of class WzCampaignSubInfoBox (campaign_title_c); the initialisers are its
' _construct's.
Option Explicit

Public leftPadding As Long                 ' = 40
Public topPadding As Long                  ' = 20
Public imgDimensions As Long               ' = 16
Public imgTextPadding As Long              ' = 5
Public pUniverseTexture As gfx_api_texture ' gfx_api::texture* = nullptr
Public universeTextureColor As Long        ' PIELIGHT = pal_RGBA(255,255,255,255)
Public universeLabel As WIDGET             ' std::shared_ptr<W_LABEL>
Public compatibilityResult As WzModCompatibilityResult  ' = WzModCompatibilityResult::OK
Public compatibilityLabel As WIDGET        ' std::shared_ptr<W_LABEL>
