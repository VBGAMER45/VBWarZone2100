VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlayerReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' netplay.h: struct PlayerReference. This structure provides read-only access to a player, and can be used
' to identify players uniquely. It holds the player data after the player has disconnected, and it is
' released automatically by reference counting (the std::shared_ptr is the object reference).
' operator-> is the PLAYER's fields as properties (name_ for name, colour, ...), of the slot or of the detached copy.
Option Explicit

Private detached As PLAYER_             ' std::unique_ptr<PLAYER> detached
Private hasDetached As Boolean          ' detached != nullptr
Private index_ As Long                  ' uint32_t index

Friend Sub construct(ByVal index As Long)
    index_ = index
End Sub

Public Sub disconnect()
    NetPlay_players_copy index_, detached           ' std::make_unique<PLAYER>(NetPlay.players[index])
    Set detached.wzFiles = New Collection           ' detached->wzFiles = std::make_shared<std::vector<WZFile>>()
    hasDetached = True
End Sub

Public Function isHost() As Boolean
    isHost = index_ = NetPlay.hostPlayer
End Function

Public Function isDetached() As Boolean
    isDetached = hasDetached
End Function

Public Function originalIndex() As Long
    originalIndex = index_
End Function

' PLAYER const *operator ->() const: the fields
Private Sub current(p As PLAYER_)
    If hasDetached Then p = detached Else NetPlay_players_copy index_, p
End Sub

Public Function name_() As String
    Dim p As PLAYER_
    current p
    name_ = p.name
End Function

Public Function position() As Long
    Dim p As PLAYER_
    current p
    position = p.position
End Function

Public Function colour() As Long
    Dim p As PLAYER_
    current p
    colour = p.colour
End Function

Public Function allocated() As Boolean
    Dim p As PLAYER_
    current p
    allocated = p.allocated
End Function

Public Function team() As Long
    Dim p As PLAYER_
    current p
    team = p.team
End Function

Public Function isSpectator() As Boolean
    Dim p As PLAYER_
    current p
    isSpectator = p.isSpectator
End Function

Public Function isAdmin() As Boolean
    Dim p As PLAYER_
    current p
    isAdmin = p.isAdmin
End Function

Public Function difficulty() As Long
    Dim p As PLAYER_
    current p
    difficulty = p.difficulty
End Function

Public Function IPtextAddress() As String
    Dim p As PLAYER_
    current p
    IPtextAddress = p.IPtextAddress
End Function
