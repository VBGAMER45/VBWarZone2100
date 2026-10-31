VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzPlayerMuteStatusLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzPlayerMuteStatusLabel
' src/hci/chatoptions.cpp: the fields of class WzPlayerMuteStatusLabel (chatoptions_c); optional<PIELIGHT> muteIconColorOverride is has_muteIconColorOverride and the value; the initialisers are its _construct's.
Option Explicit

Public playerIdx As Long                   ' uint32_t = 0
Public FontID As iV_fonts                  ' = font_regular
Public pText As String                     ' WzString
Public wzText As WzText
Public has_muteIconColorOverride As Boolean ' optional<PIELIGHT> muteIconColorOverride
Public muteIconColorOverride As Long
Public isSelectedPlayer As Boolean         ' = false
Public isDisabled As Boolean               ' = false

Private Sub Class_Initialize()
    Set wzText = New WzText
End Sub
