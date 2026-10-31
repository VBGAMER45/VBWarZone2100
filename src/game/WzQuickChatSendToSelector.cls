VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzQuickChatSendToSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzQuickChatSendToSelector
' src/hci/quickchat.cpp: the fields of class WzQuickChatSendToSelector (quickchat_c); the UDT fields are Friend properties (a copy in, a copy out); weakSerial: (VB) its std::weak_ptrs; the initialisers are its _construct's.
Option Explicit

Public weakSerial As Long
Private playersToShow_ As UInt32Vector     ' std::vector<uint32_t>
Private playerTeammates_ As UInt32Vector
Private playerHumanTeammates_ As UInt32Vector
Private playerAITeammates_ As UInt32Vector
Private playerSelections_() As Long        ' std::array<PlayerSelectionState, MAX_CONNECTED_PLAYERS>
Private targeting_ As WzQuickChatTargeting
Public playersSelectorWidget As WIDGET     ' std::shared_ptr<WzQuickChatSendToPlayersSelector>
Public overlayScreen As W_SCREEN           ' std::shared_ptr<W_SCREEN>
Public sendToText As WzText
Public cachedPlayerText As WzText
Public teamOnly As Boolean                 ' = false
Public hasHighlight As Boolean             ' = false

Private Sub Class_Initialize()
    ReDim playerSelections_(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1)
    Set sendToText = New WzText
    Set cachedPlayerText = New WzText
End Sub

Friend Property Get playersToShow() As UInt32Vector
    playersToShow = playersToShow_
End Property

Friend Property Let playersToShow(v As UInt32Vector)
    playersToShow_ = v
End Property

Friend Property Get playerTeammates() As UInt32Vector
    playerTeammates = playerTeammates_
End Property

Friend Property Let playerTeammates(v As UInt32Vector)
    playerTeammates_ = v
End Property

Friend Property Get playerHumanTeammates() As UInt32Vector
    playerHumanTeammates = playerHumanTeammates_
End Property

Friend Property Let playerHumanTeammates(v As UInt32Vector)
    playerHumanTeammates_ = v
End Property

Friend Property Get playerAITeammates() As UInt32Vector
    playerAITeammates = playerAITeammates_
End Property

Friend Property Let playerAITeammates(v As UInt32Vector)
    playerAITeammates_ = v
End Property

Public Property Get playerSelections(ByVal i As Long) As WzQuickChatSendToSelector_PlayerSelectionState
    playerSelections = playerSelections_(i)
End Property

Public Property Let playerSelections(ByVal i As Long, ByVal v As WzQuickChatSendToSelector_PlayerSelectionState)
    playerSelections_(i) = v
End Property

Friend Property Get targeting() As WzQuickChatTargeting
    targeting = targeting_
End Property

Friend Property Let targeting(v As WzQuickChatTargeting)
    targeting_ = v
End Property
