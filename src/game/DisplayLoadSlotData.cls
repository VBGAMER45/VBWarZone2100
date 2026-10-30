VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisplayLoadSlotData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/challenge.cpp: struct DisplayLoadSlotData (with its struct DisplayLoadSlotCache cache), the pUserData of
' the challenge menu's slots: the button's text as drawn and the challenge's file. A new one is C++'s
' new DisplayLoadSlotData(); the widget's onDelete drops it (pUserData = Empty).
Option Explicit

Public cache_fullText As String         ' cache.fullText
Public cache_wzText As WzText           ' cache.wzText
Public filename As String               ' const char *: vbNullString = nullptr

Private Sub Class_Initialize()
    Set cache_wzText = New WzText
    filename = vbNullString
End Sub
