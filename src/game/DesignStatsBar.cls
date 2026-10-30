VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DesignStatsBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class DesignStatsBar
' src/design.cpp: DesignStatsBar's fields (design_c); the member initialisers are DesignStatsBar_construct's.
' Animation<float> minorAnimation is a Private animation_c UDT with Friend accessors.
Option Explicit

Public valueText As WzText
Public deltaText As WzText
Public maxValueTextWidth As Long        ' uint32_t = iV_GetTextWidth("00000", font_regular)
Private m_minorAnimation As Animation_float ' = Animation<float>(realTime, EASE_IN, 200).setInitialData(0).setFinalData(0)
Public PADDING_ As Long                 ' const uint32_t = 3
Public lessIsBetter As Boolean          ' = false

Friend Property Get minorAnimation() As Animation_float
    minorAnimation = m_minorAnimation
End Property

Friend Property Let minorAnimation(v As Animation_float)
    m_minorAnimation = v
End Property

Private Sub Class_Initialize()
    Set valueText = New WzText
    Set deltaText = New WzText
End Sub
