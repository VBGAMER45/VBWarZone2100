VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LobbyGameModsWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class LobbyGameModsWidget
' src/titleui/widgets/gamebrowserform.cpp: the fields of class LobbyGameModsWidget (gamebrowserform_c)
Option Explicit

Public modsListWidg As WIDGET           ' std::shared_ptr<W_LABEL>
Public limitWidgets As Collection       ' std::vector<std::shared_ptr<WIDGET>>
Public limitsListWidg As WIDGET         ' std::shared_ptr<W_LABEL>

Private Sub Class_Initialize()
    Set limitWidgets = New Collection
End Sub
