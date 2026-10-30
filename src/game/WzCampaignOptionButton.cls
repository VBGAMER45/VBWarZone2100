VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCampaignOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzCampaignOptionButton
' src/titleui/campaign.cpp: the fields of class WzCampaignOptionButton (campaign_title_c); the initialisers are its
' _construct's.
Option Explicit

Public wzText As WzText
Public wzDescriptionText As WzText
Public descriptionFontID As iV_fonts       ' = font_small
Public cachedIdealWidth As Long            ' int32_t = 0
Public lastWidgetWidth As Long             ' = 0
Public horizontalPadding As Long           ' = 10
Public verticalPadding As Long             ' = 5
Public internalPadding As Long             ' = 3

Private Sub Class_Initialize()
    Set wzText = New WzText
    Set wzDescriptionText = New WzText
End Sub
