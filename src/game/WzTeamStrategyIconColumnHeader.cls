VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzTeamStrategyIconColumnHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzTeamStrategyIconColumnHeader
' src/hci/teamstrategy.cpp: the fields of class WzTeamStrategyIconColumnHeader (teamstrategy_c); the initialisers are its _construct's.
Option Explicit

' columnImageRef (a UDT, teamstrategy_c's WzTeamStrategyColumnImageRef) is a Friend property
Public columnText As WzText
Public pTip As String                   ' std::string
Public disabledColumn As Boolean        ' = false

Private columnImageRef_ As WzTeamStrategyColumnImageRef

Friend Property Get columnImageRef() As WzTeamStrategyColumnImageRef
    columnImageRef = columnImageRef_
End Property

Friend Property Let columnImageRef(v As WzTeamStrategyColumnImageRef)
    columnImageRef_ = v
End Property

Private Sub Class_Initialize()
    Set columnText = New WzText
End Sub
