Attribute VB_Name = "stubs_net"
' Stand-ins for lib/netplay until its cluster ports it, with the C++ names
' and signatures: netplay.cpp's settings (kept as in C++; nothing connects),
' the NETPLAY fields read so far, the command line's join target,
' SpectatorInfo::currentNetPlayState and netlobby.h's GetDefaultLobbyAddress.
' netplay.cpp's message traffic and NET_InitPlayer are netplay_c.
'
' std::vector<PLAYER> NetPlay.players: one array per field, with room for
' MAX_CONNECTED_PLAYERS + 1 players (the replay spectator's slot, which
' restoreOptions pushes back) and its size in playersSize (NETPLAY's
' constructor: MAX_CONNECTED_PLAYERS). PLAYER() (push_back, resize) is the
' value-initialised struct: every field 0, false or "".
Option Explicit

' nettypes.h
Public Const MAX_SPECTATOR_SLOTS As Long = 10
Public Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' netplay.h: struct NETPLAY (the fields used so far). The type gets the
' underscore: the variable NetPlay is used in far more places.
Public Type NETPLAY_
    bComms As Boolean                   ' Actually do the comms?
    isHost As Boolean                   ' True if we are hosting the game
    isPortMappingEnabled As Boolean     ' if we want the automatic Port mapping setup routines to run
    hostPlayer As Long                  ' uint32_t: Index of host in player array (NET_InitPlayers: NET_HOST_ONLY, 0)
    isHostAlive As Boolean              ' if the host is still alive
End Type

Public NetPlay As NETPLAY_
Private playerAllocated(0 To MAX_CONNECTED_PLAYERS) As Boolean   ' NetPlay.players[i].allocated: Allocated as a human player
Private playerIsSpectator(0 To MAX_CONNECTED_PLAYERS) As Boolean ' NetPlay.players[i].isSpectator: whether this slot is a spectator slot
Private playerFaction(0 To MAX_CONNECTED_PLAYERS) As Long       ' NetPlay.players[i].faction (FactionID; FACTION_NORMAL, 0)
Private playerPosition(0 To MAX_CONNECTED_PLAYERS) As Long      ' NetPlay.players[i].position: Map starting position (the first players: -1)
Private playerDifficulty(0 To MAX_CONNECTED_PLAYERS) As Long    ' NetPlay.players[i].difficulty (AIDifficulty; the first players: DISABLED, -1)
Private playerName(0 To MAX_CONNECTED_PLAYERS) As String        ' NetPlay.players[i].name (char[StringSize]: Player name)
Private playerIsAdmin(0 To MAX_CONNECTED_PLAYERS) As Boolean    ' NetPlay.players[i].isAdmin: whether this slot has admin privs
Private playerColour(0 To MAX_CONNECTED_PLAYERS) As Long        ' NetPlay.players[i].colour: Which colour slot this player is using
Private playerHeartattacktime(0 To MAX_CONNECTED_PLAYERS) As Long   ' uint32_t NetPlay.players[i].heartattacktime: Time cardiac arrest started
Private playerHeartbeat(0 To MAX_CONNECTED_PLAYERS) As Boolean  ' NetPlay.players[i].heartbeat: If connection is still alive or not
Private playerTeam(0 To MAX_CONNECTED_PLAYERS) As Long          ' NetPlay.players[i].team: Which team we are on (int32_t::max for spectator team)
Private playerReady(0 To MAX_CONNECTED_PLAYERS) As Boolean      ' NetPlay.players[i].ready: player ready to start?
Private playerAi(0 To MAX_CONNECTED_PLAYERS) As Long            ' int8_t NetPlay.players[i].ai: index into sorted list of AIs, zero is always default AI
Private playerAutoGame(0 To MAX_CONNECTED_PLAYERS) As Boolean   ' NetPlay.players[i].autoGame: if we are running a autogame (AI controls us)
Private playersSize As Long             ' NetPlay.players.size()

Public netGameserverPortOverride As Boolean
Public iptoconnect As String            ' char[PATH_MAX]: holds IP/hostname from command line
Public cliConnectAsSpectator As Boolean ' for cli option
Private bAsyncJoinApprovalEnabled As Boolean

Private masterserver_name As String
Private gameserver_port As Long         ' unsigned int
Private bJoinPrefTryIPv6First As Boolean
Private bDefaultHostFreeChatEnabled As Boolean
Private bEnableTCPNoDelay As Boolean
Private hostJoinOptions_botProt As Long         ' netlobby::HostJoinOptions
Private hostJoinOptions_proxyIPs As Long
Private hostJoinOptions_hostingIPs As Long
Private netStubsReady As Boolean

' the C++ initialisers
Private Sub InitNetStubs()
    Dim i As Long
    If netStubsReady Then Exit Sub
    netStubsReady = True
    playersSize = MAX_CONNECTED_PLAYERS     ' NETPLAY(): players.resize(MAX_CONNECTED_PLAYERS), value-initialised PLAYERs (all 0)
    bJoinPrefTryIPv6First = True
    bDefaultHostFreeChatEnabled = True
    bEnableTCPNoDelay = True
    hostJoinOptions_botProt = 0         ' HostJoinOptionValue::BlockAll
    hostJoinOptions_proxyIPs = 1        ' HostJoinOptionValue::BlockSuspicious
    hostJoinOptions_hostingIPs = 0      ' HostJoinOptionValue::BlockAll
End Sub

' netlobby.h: lobbyAddressEncoded, each byte xor scr_key (0x5A)
Public Function netlobby_GetDefaultLobbyAddress() As String
    Dim lobbyAddressEncoded As Variant, i As Long, s As String
    lobbyAddressEncoded = Array(&H32, &H2E, &H2E, &H2A, &H29, &H60, &H75, &H75, &H2D, &H20, &H36, &H35, &H38, &H38, &H23, &H74, _
        &H2D, &H20, &H68, &H6B, &H6A, &H6A, &H74, &H34, &H3F, &H2E, &H75, &H36, &H35, &H38, &H38, &H23)
    For i = 0 To UBound(lobbyAddressEncoded)
        s = s & ChrW$(lobbyAddressEncoded(i) Xor &H5A)
    Next
    netlobby_GetDefaultLobbyAddress = s
End Function

Public Sub NETsetLobbyserverAddress(ByVal lobbyAddress As String)
    masterserver_name = lobbyAddress
End Sub

Public Function NETgetLobbyserverAddress() As String
    NETgetLobbyserverAddress = masterserver_name
End Function

Public Sub NETsetGameserverPort(ByVal port As Long)
    gameserver_port = port
End Sub

Public Function NETgetGameserverPort() As Long
    NETgetGameserverPort = gameserver_port
End Function

Public Sub NETsetJoinPreferenceIPv6(ByVal bTryIPv6First As Boolean)
    InitNetStubs
    bJoinPrefTryIPv6First = bTryIPv6First
End Sub

Public Function NETgetJoinPreferenceIPv6() As Boolean
    InitNetStubs
    NETgetJoinPreferenceIPv6 = bJoinPrefTryIPv6First
End Function

Public Sub NETsetDefaultMPHostFreeChatPreference(ByVal enabled As Boolean)
    InitNetStubs
    bDefaultHostFreeChatEnabled = enabled
End Sub

Public Function NETgetDefaultMPHostFreeChatPreference() As Boolean
    InitNetStubs
    NETgetDefaultMPHostFreeChatPreference = bDefaultHostFreeChatEnabled
End Function

Public Sub NETsetEnableTCPNoDelay(ByVal enabled As Boolean)
    InitNetStubs
    bEnableTCPNoDelay = enabled
End Sub

Public Function NETgetEnableTCPNoDelay() As Boolean
    InitNetStubs
    NETgetEnableTCPNoDelay = bEnableTCPNoDelay
End Function

Public Sub NETgetLobbyHostJoinOptions(botProt As Long, proxyIPs As Long, hostingIPs As Long)
    InitNetStubs
    botProt = hostJoinOptions_botProt
    proxyIPs = hostJoinOptions_proxyIPs
    hostingIPs = hostJoinOptions_hostingIPs
End Sub

' C: also NETupdateLobbyGameListing() (no game is listed yet)
Public Sub NETsetLobbyHostJoinOptions(ByVal botProt As Long, ByVal proxyIPs As Long, ByVal hostingIPs As Long)
    InitNetStubs
    hostJoinOptions_botProt = botProt
    hostJoinOptions_proxyIPs = proxyIPs
    hostJoinOptions_hostingIPs = hostingIPs
End Sub

' C: counts the spectator slots of NetPlay.players (none yet); here
' SpectatorInfo::toUint32() of the result
Public Function SpectatorInfo_currentNetPlayState() As Long
    SpectatorInfo_currentNetPlayState = 0
End Function

Public Function NetPlay_players_allocated(ByVal i As Long) As Boolean
    NetPlay_players_allocated = playerAllocated(i)
End Function

Public Sub set_NetPlay_players_allocated(ByVal i As Long, ByVal value As Boolean)
    playerAllocated(i) = value
End Sub

Public Function NetPlay_players_isSpectator(ByVal i As Long) As Boolean
    NetPlay_players_isSpectator = playerIsSpectator(i)
End Function

Public Sub set_NetPlay_players_isSpectator(ByVal i As Long, ByVal value As Boolean)
    playerIsSpectator(i) = value
End Sub

Public Function NetPlay_players_faction(ByVal i As Long) As Long
    NetPlay_players_faction = playerFaction(i)
End Function

Public Sub set_NetPlay_players_faction(ByVal i As Long, ByVal value As Long)
    playerFaction(i) = value And &HFF&
End Sub

Public Function NetPlay_players_position(ByVal i As Long) As Long
    InitNetStubs
    NetPlay_players_position = playerPosition(i)
End Function

Public Sub set_NetPlay_players_position(ByVal i As Long, ByVal value As Long)
    InitNetStubs
    playerPosition(i) = value
End Sub

Public Function NetPlay_players_difficulty(ByVal i As Long) As Long
    InitNetStubs
    NetPlay_players_difficulty = playerDifficulty(i)
End Function

Public Sub set_NetPlay_players_difficulty(ByVal i As Long, ByVal value As Long)
    InitNetStubs
    playerDifficulty(i) = value
End Sub

Public Function NetPlay_players_name(ByVal i As Long) As String
    NetPlay_players_name = playerName(i)
End Function

Public Sub set_NetPlay_players_name(ByVal i As Long, ByVal value As String)
    playerName(i) = value
End Sub

Public Function NetPlay_players_isAdmin(ByVal i As Long) As Boolean
    NetPlay_players_isAdmin = playerIsAdmin(i)
End Function

Public Sub set_NetPlay_players_isAdmin(ByVal i As Long, ByVal value As Boolean)
    playerIsAdmin(i) = value
End Sub

Public Function NetPlay_players_colour(ByVal i As Long) As Long
    NetPlay_players_colour = playerColour(i)
End Function

Public Sub set_NetPlay_players_colour(ByVal i As Long, ByVal value As Long)
    playerColour(i) = value
End Sub

Public Function NetPlay_players_heartattacktime(ByVal i As Long) As Long
    NetPlay_players_heartattacktime = playerHeartattacktime(i)
End Function

Public Sub set_NetPlay_players_heartattacktime(ByVal i As Long, ByVal value As Long)
    playerHeartattacktime(i) = value
End Sub

Public Function NetPlay_players_heartbeat(ByVal i As Long) As Boolean
    NetPlay_players_heartbeat = playerHeartbeat(i)
End Function

Public Sub set_NetPlay_players_heartbeat(ByVal i As Long, ByVal value As Boolean)
    playerHeartbeat(i) = value
End Sub

Public Function NetPlay_players_team(ByVal i As Long) As Long
    NetPlay_players_team = playerTeam(i)
End Function

Public Sub set_NetPlay_players_team(ByVal i As Long, ByVal value As Long)
    playerTeam(i) = value
End Sub

Public Function NetPlay_players_ready(ByVal i As Long) As Boolean
    NetPlay_players_ready = playerReady(i)
End Function

Public Sub set_NetPlay_players_ready(ByVal i As Long, ByVal value As Boolean)
    playerReady(i) = value
End Sub

' int8_t: -128 .. 127
Public Function NetPlay_players_ai(ByVal i As Long) As Long
    NetPlay_players_ai = playerAi(i)
End Function

Public Sub set_NetPlay_players_ai(ByVal i As Long, ByVal value As Long)
    value = value And &HFF&
    If value >= 128 Then value = value - 256
    playerAi(i) = value
End Sub

Public Function NetPlay_players_autoGame(ByVal i As Long) As Boolean
    NetPlay_players_autoGame = playerAutoGame(i)
End Function

Public Sub set_NetPlay_players_autoGame(ByVal i As Long, ByVal value As Boolean)
    playerAutoGame(i) = value
End Sub

' NetPlay.players.size()
Public Function NetPlay_players_size() As Long
    InitNetStubs
    NetPlay_players_size = playersSize
End Function

' (VB) players[i] = PLAYER(): the value-initialised struct
Private Sub PLAYER_construct(ByVal i As Long)
    playerName(i) = ""
    playerPosition(i) = 0
    playerColour(i) = 0
    playerAllocated(i) = False
    playerHeartattacktime(i) = 0
    playerHeartbeat(i) = False
    playerTeam(i) = 0
    playerReady(i) = False
    playerAi(i) = 0
    playerDifficulty(i) = 0
    playerAutoGame(i) = False
    playerFaction(i) = 0
    playerIsSpectator(i) = False
    playerIsAdmin(i) = False
End Sub

' NetPlay.players.push_back(PLAYER())
Public Sub NetPlay_players_push_back()
    InitNetStubs
    If Not (playersSize <= MAX_CONNECTED_PLAYERS) Then
        ' (VB) a std::vector grows without a limit; here the replay spectator's slot is the last there is
        debug_ 0, LOG_ERROR, "NetPlay_players_push_back", "(VB) No room for another player (%u)", playersSize
        Exit Sub
    End If
    PLAYER_construct playersSize
    playersSize = playersSize + 1
End Sub

' NetPlay.players.resize(n): new players are PLAYER()
Public Sub NetPlay_players_resize(ByVal n As Long)
    InitNetStubs
    If Not (n >= 0 And n <= MAX_CONNECTED_PLAYERS + 1) Then
        debug_ 0, LOG_ERROR, "NetPlay_players_resize", "(VB) No room for %u players", n
        Exit Sub
    End If
    Do While playersSize < n
        PLAYER_construct playersSize
        playersSize = playersSize + 1
    Loop
    playersSize = n
End Sub

Public Sub NETsetAsyncJoinApprovalRequired(ByVal enabled As Boolean)
    bAsyncJoinApprovalEnabled = enabled
End Sub

'--------------------------------------------------------------------------
' cluster A2: netplay.cpp for init.cpp's systemShutdown (no game is open)

Public Function netplayShutDown() As Boolean
    netplayShutDown = True
End Function
