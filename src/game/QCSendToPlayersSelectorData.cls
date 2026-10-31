VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QCSendToPlayersSelectorData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzQuickChatSendToPlayersSelector
' src/hci/quickchat.cpp: the fields of class WzQuickChatSendToPlayersSelector (quickchat_c) (VB's 31-character limit: the class's name is too long for its data class); playerButtons is a Collection (element n is item n + 1); the initialisers are its _construct's.
Option Explicit

Public humanTeammates As WIDGET            ' std::shared_ptr<WzQuickChatToggleMixedCheckButton>
Public aiTeammates As WIDGET               ' std::shared_ptr<WzQuickChatToggleMixedCheckButton>
Public playerButtons As Collection         ' std::vector<std::shared_ptr<WzQuickChatPlayerButton>>
Public playersList As WIDGET               ' std::shared_ptr<ScrollableListWidget>

Private Sub Class_Initialize()
    Set playerButtons = New Collection
End Sub
