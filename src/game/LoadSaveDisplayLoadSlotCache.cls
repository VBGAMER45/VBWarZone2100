VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadSaveDisplayLoadSlotCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/loadsave.cpp: struct LoadSaveDisplayLoadSlotCache, the pUserData of the load/save screen's slots: the
' button's text as drawn. A new one is C++'s new LoadSaveDisplayLoadSlotCache(); the widget's onDelete drops it
' (pUserData = Empty). (VB) wzText is wzText_ (a field named as its class would recase WzText).
Option Explicit

Public fullText As String               ' std::string
Public wzText_ As WzText                ' WzText wzText

Private Sub Class_Initialize()
    Set wzText_ = New WzText
End Sub
