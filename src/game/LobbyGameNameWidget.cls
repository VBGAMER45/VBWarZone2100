VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LobbyGameNameWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class LobbyGameNameWidget
' src/titleui/widgets/gamebrowserform.cpp: the fields of class LobbyGameNameWidget (gamebrowserform_c)
Option Explicit

Public roomNameWidg As WIDGET           ' std::shared_ptr<W_LABEL>
Public hostNameWidg As WIDGET           ' std::shared_ptr<W_LABEL>
