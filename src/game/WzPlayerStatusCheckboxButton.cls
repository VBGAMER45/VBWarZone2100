VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzPlayerStatusCheckboxButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzPlayerStatusCheckboxButton
' src/screens/netpregamescreen.cpp: the fields of class WzPlayerStatusCheckboxButton (netpregamescreen_c); the initialisers are its _construct's.
Option Explicit

Public status As Long                   ' enum class LoadingStatus = LoadingStatus::Loading (0)
Public player As Long                   ' uint32_t
Public wzText As WzText
Public checkmarkText As WzText          ' std::shared_ptr<WzText>
Public optionsOverlayScreen As W_SCREEN ' std::shared_ptr<W_SCREEN>

Private Sub Class_Initialize()
    Set wzText = New WzText
End Sub
