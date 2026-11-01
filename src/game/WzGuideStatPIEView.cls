VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGuideStatPIEView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzGuideStatPIEView
' src/screens/guidescreen.cpp: WzGuideStatPIEView's fields (guidescreen_c); the initialisers are WzGuideStatPIEView_construct's
' and Class_Initialize's.
Option Explicit

Public Stat As Long                          ' BASE_STATS * = nullptr
Public design As Long                        ' DROID_TEMPLATE design: (VB) its template_c pool slot, made by makeFromDroidTemplateJSON (0: none)
