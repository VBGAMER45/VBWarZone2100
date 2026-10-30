VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class ObjectsForm
' src/hci/objects_stats.h: ObjectsForm's fields (objects_stats_c)
Option Explicit

Public objectsList As WIDGET            ' std::shared_ptr<IntListTabWidget>
Public buttonsCount As Long             ' size_t = 0
Public previousHighlighted As Long      ' BASE_OBJECT * = nullptr
