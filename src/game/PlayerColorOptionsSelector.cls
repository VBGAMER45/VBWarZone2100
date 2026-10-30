VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlayerColorOptionsSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/options/defaultsoptions.cpp: the fields of class PlayerColorOptionsSelector
' (VB) data of widget class PlayerColorOptionsSelector
Option Explicit

Public getCurrentValueFunc As WzFunc     ' GetCurrentValueFunc: optional<unsigned>() (a Variant: Empty is nullopt)
Public onSelectFunc As WzFunc            ' OnSelectColorFunc: bool (optional<unsigned> tc)
Public onChangeHandlers As Collection    ' std::vector<std::function<void(WIDGET&)>>
Public colorButtons As Collection        ' std::vector<std::shared_ptr<WzMultiButton>>
Public buttonWidth As Long               ' = 0
Public buttonHeight As Long              ' = 0
Public buttonMargin As Long              ' = 5
