VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzClippableButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/options/optionsforms.cpp: the fields of class WzClippableButton
' (VB) data of widget class WzClippableButton
Option Explicit

Public wzText As WzCachedText
Public horizontalPadding As Long         ' = 15
Public verticalPadding As Long           ' = 10
