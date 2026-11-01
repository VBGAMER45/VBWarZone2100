VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGuideRetButtonView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzGuideRetButtonView
' src/screens/guidescreen.cpp: WzGuideRetButtonView's fields (guidescreen_c); the initialisers are WzGuideRetButtonView_construct's
' and Class_Initialize's.
Option Explicit

Public retButtonDisplayFilename As String    ' WzString
Public imageWidth As Long                    ' int32_t = 0
Public imageHeight As Long                   ' int32_t = 0
