VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzThrottledUpdateLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzThrottledUpdateLabel
' src/spectatorwidgets.cpp: the fields of class WzThrottledUpdateLabel (spectatorwidgets_c); the initialisers are its _construct's.
Option Explicit

Public updateFunc As WzFunc             ' std::function<void (W_LABEL& label)>
Public updateInterval As Long           ' uint32_t = 0
Public lastGameTimeUpdated As Long      ' uint32_t = 0
