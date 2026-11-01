VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CachedConsoleMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/screens/netpregamescreen.cpp: WzGameStartOverlayScreen_CLICKFORM's struct CachedConsoleMessage (netpregamescreen_c)
Option Explicit

Public display As WzText
Public timeAdded As Currency            ' std::chrono::steady_clock::time_point: a QueryPerformanceCounter count
Public duration As Long                 ' std::chrono::milliseconds
Public JustifyType As CONSOLE_TEXT_JUSTIFICATION  ' text justification
Public player As Long                   ' Player who sent this message or SYSTEM_MESSAGE

Private Sub Class_Initialize()
    Set display = New WzText
End Sub
