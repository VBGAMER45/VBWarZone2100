VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzTeamStrategyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzTeamStrategyButton
' src/hci/teamstrategy.cpp: the fields of class WzTeamStrategyButton (teamstrategy_c); the initialisers are its _construct's.
Option Explicit

' imageRef (a UDT, teamstrategy_c's WzTeamStrategyColumnImageRef) is a Friend property
Public parentStrategyView As Long       ' std::weak_ptr<TeamStrategyView>: its address (WIDGET_fromPtr)
Public sharedImageManager As WzTeamStratImagesManager  ' std::shared_ptr<WzTeamStrategyColumnImagesManager>
Public displayWidth As Long
Public displayHeight As Long

Private imageRef_ As WzTeamStrategyColumnImageRef

Friend Property Get imageRef() As WzTeamStrategyColumnImageRef
    imageRef = imageRef_
End Property

Friend Property Let imageRef(v As WzTeamStrategyColumnImageRef)
    imageRef_ = v
End Property
