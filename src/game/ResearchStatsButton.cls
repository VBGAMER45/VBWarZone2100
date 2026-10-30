VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResearchStatsButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class ResearchStatsButton
' src/hci/research.cpp: ResearchStatsButton's fields (research_ui_c)
Option Explicit

Public controller As ResearchController         ' std::shared_ptr<ResearchController>
Public objectIndex As Long                      ' size_t
