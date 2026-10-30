VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WidgetInfoButtonWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WidgetInfoButtonWrapper
' src/titleui/campaign.cpp: the fields of class WidgetInfoButtonWrapper (campaign_title_c); the initialisers are its
' _construct's.
Option Explicit

Public widget As WIDGET                    ' std::shared_ptr<WIDGET>
Public infoButton As WIDGET                ' std::shared_ptr<WzInfoButton>
Public infoClickHandler As WzFunc          ' InfoClickHandler (Nothing: empty)
Public topPadding As Long                  ' = 0
Public bottomPadding As Long               ' = 0
Public internalHorizontalPadding As Long   ' = 10
Public imageDimensions As Long             ' const int = 16
