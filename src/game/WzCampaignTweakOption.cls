VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCampaignTweakOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/modinfo.h: struct WzCampaignTweakOption (modinfo.cpp, which reads them, is not ported). A
' WzJsonLocalizedString is its getLocalizedString().
Option Explicit

Public uniqueIdentifier As String
Public type_ As WzCampaignTweakOption_Type
Public enabled As Boolean                  ' = false
Public userEditable As Boolean             ' = true
Public displayName As String               ' WzJsonLocalizedString
Public description As String               ' WzJsonLocalizedString

Private Sub Class_Initialize()
    userEditable = True
End Sub
