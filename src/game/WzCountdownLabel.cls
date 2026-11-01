VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCountdownLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzCountdownLabel
' src/screens/netpregamescreen.cpp: the fields of class WzCountdownLabel (netpregamescreen_c); the initialisers are its _construct's.
Option Explicit

Public has_startTime As Boolean         ' optional<std::chrono::steady_clock::time_point> startTime: has_value()
Public startTime As Currency            ' its value: a QueryPerformanceCounter count
Public countdownSeconds As Double       ' std::chrono::seconds::rep = 0
Public maxWidth As Long                 ' int32_t = 0
Public FontID As iV_fonts               ' = font_large
Public fontColour As Long               ' PIELIGHT = pal_RGBA(255, 255, 255, 255)
Public displayedNumber As WzText

Private Sub Class_Initialize()
    FontID = font_large
    fontColour = pal_RGBA(255, 255, 255, 255)
    Set displayedNumber = New WzText
End Sub
