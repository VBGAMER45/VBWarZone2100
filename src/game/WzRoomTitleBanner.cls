VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzRoomTitleBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzRoomTitleBanner
' src/titleui/widgets/blindwaitingroom.cpp: the fields of class WzRoomTitleBanner (blindwaitingroom_c); the initialisers are its _construct's.
Option Explicit

Public infoClickHandler As WzFunc       ' InfoClickHandler (std::function<void ()>)
Public topPadding As Long               ' = 10
Public bottomPadding As Long            ' = 10
Public internalHorizontalPadding As Long ' = 10
Public titleLabel As WIDGET             ' std::shared_ptr<W_LABEL>
Public infoButton As WIDGET             ' std::shared_ptr<WzInfoButton>
