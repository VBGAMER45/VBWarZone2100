VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzPlayerBlindWaitingRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzPlayerBlindWaitingRoom
' src/titleui/widgets/blindwaitingroom.cpp: the fields of class WzPlayerBlindWaitingRoom (blindwaitingroom_c); the initialisers are its _construct's.
Option Explicit

' (std::weak_ptr<WzMultiplayerOptionsTitleUI> weakTitleUI: never read, not kept)
Public titleBannerBottomPadding As Long ' = 10
Public outerWidgetPaddingX As Long      ' = 15
Public betweenWidgetPaddingY As Long    ' = 10
Public bottomPaddingY As Long           ' = 15
Public titleBanner As WIDGET            ' std::shared_ptr<WzRoomTitleBanner>
Public hostInfo As WIDGET               ' std::shared_ptr<WzBlindRoomHostInfo>
Public playerInfo As WIDGET             ' std::shared_ptr<WzBlindRoomPlayerInfo>
Public statusInfo As WIDGET             ' std::shared_ptr<WzBlindRoomStatusInfo>
Public playersReadyInfo As WIDGET       ' std::shared_ptr<WzBlindPlayersReadyInfo>
