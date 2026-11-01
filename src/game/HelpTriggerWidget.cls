VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HelpTriggerWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class HelpTriggerWidget
' src/screens/helpscreen.cpp: the fields of class HelpTriggerWidget (helpscreen_c); the initialisers are its _construct's.
Option Explicit

Public parent As Long                   ' std::weak_ptr<W_HELPSCREEN_CLICKFORM>: its widgweak_vb serial
Public target As WIDGET                 ' std::weak_ptr<WIDGET>: (VB) held strongly (helpscreen_c's header)
Public has_hoverStart As Boolean        ' optional<std::chrono::steady_clock::time_point> hoverStart
Public hoverStart As Double             ' (widgSteadyNowMs)
Public lastUpdatedIndicatorAnimationTime As Long ' uint32_t = 0
Public indicatorAnimationProgress As Single ' float = 0.f
Public indicator_inset_x As Long        ' Vector2i indicator_inset = Vector2i(0, 0)
Public indicator_inset_y As Long
Public clickedDownOnThis As Boolean     ' = false
Public queuedWidgetShowOnRelease As Boolean ' = false
Public hasFocus As Boolean              ' = false
Public isHighlight As Boolean           ' = false
Public weakSerial As Long               ' (VB) its widgweak_vb serial (its weak pointers)
