VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCampaignHelpPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzCampaignHelpPanel
' src/titleui/campaign.cpp: the fields of class WzCampaignHelpPanel (campaign_title_c); the initialisers are its
' _construct's.
Option Explicit

Public formTitle As WIDGET                 ' std::shared_ptr<W_LABEL>
Public closeButton As WIDGET               ' std::shared_ptr<WzFrontendImageButton>
Public paragraphContainer As WIDGET        ' std::shared_ptr<ScrollableListWidget>
Public baseFontId As iV_fonts              ' = font_regular
Public innerPadding As Long                ' = 20
Public betweenButtonPadding As Long        ' = 15
