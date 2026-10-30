VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoopProductionLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class LoopProductionLabel
' src/hci/manufacture.cpp (ManufactureStatsForm::LoopProductionLabel): LoopProductionLabel's fields (manufacture_c)
Option Explicit

Public lastProductionLoop As Long           ' uint8_t = 0
Public controller As ManufactureController  ' std::shared_ptr<ManufactureController>
