VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LobbyHostOptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/lobbyplayerboxtabs.cpp: the fields of class LobbyHostOptionsForm
' (VB) data of widget class LobbyHostOptionsForm
Option Explicit

Public weakTitleUI As Long              ' std::weak_ptr<WzMultiplayerOptionsTitleUI>: its address
Public formTitle As WIDGET              ' std::shared_ptr<W_LABEL>
Public closeButton As WIDGET            ' std::shared_ptr<WzFrontendImageButton>
Public hostOptionsForm As WIDGET        ' std::shared_ptr<OptionsForm>
