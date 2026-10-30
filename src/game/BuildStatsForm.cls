VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuildStatsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class BuildStatsForm
' src/hci/build.cpp: BuildStatsForm's fields (build_c)
Option Explicit

Public controller As BuildController    ' std::shared_ptr<BuildController>
Public obsoleteButton As WIDGET         ' std::shared_ptr<MultipleChoiceButton>
Public favoriteButton As WIDGET         ' std::shared_ptr<MultipleChoiceButton>
