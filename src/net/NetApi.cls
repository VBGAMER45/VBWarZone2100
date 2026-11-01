VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetApi"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' WzNet.dll's public interface. GlobalMultiUse: the exe calls these
' procedures as if they were its own. lib/netplay is ported into this DLL's
' modules; this class only forwards to them, with the C++ names (Contract D of
' cluster J). Calls the other way go through INetGame. A UDT does not cross:
' an EcKey or a Sha256 goes as its bytes.
Option Explicit

' connection_provider_registry.h: enum class ConnectionProviderType : uint8_t (GNS_DIRECT: no GameNetworkingSockets
' backend is registered)
Public Enum ConnectionProviderType
    ConnectionProviderType_TCP_DIRECT = 0
    ConnectionProviderType_GNS_DIRECT = 1
End Enum

' netlobby.h: enum class HostJoinOptionValue
Public Enum HostJoinOptionValue
    HostJoinOptionValue_BlockAll = 0
    HostJoinOptionValue_BlockSuspicious = 1
    HostJoinOptionValue_Allow = 2
End Enum

' netlobby.h: enum class HostJoinOptionType
Public Enum HostJoinOptionType
    HostJoinOptionType_BotProt = 0
    HostJoinOptionType_ProxyIPs = 1
    HostJoinOptionType_HostingIPs = 2
End Enum

' nettypes.h: enum QueueType
Public Enum QueueType
    QUEUE_TMP = 0
    QUEUE_NET = 1
    QUEUE_GAME = 2
    QUEUE_GAME_FORCED = 3
    QUEUE_BROADCAST = 4
    QUEUE_TRANSIENT_JOIN = 5
End Enum

' nettypes.h, netplay.h: the queues' constants
Public Enum nettypes_h_consts
    NET_NO_EXCLUDE = 255
    NET_ALL_PLAYERS = 255
    NET_HOST_ONLY = 0
    MAX_GAMEQUEUE_SLOTS = 24        ' (MAX_CONNECTED_PLAYERS + 1) /// < +1 for the replay spectator slot
    MAX_TMP_SOCKETS = 16
End Enum

' netplay.h: #define SYNC_FLAG
Public Enum netplay_h_consts
    SYNC_FLAG = &H10000000          ' special flag used for logging.
End Enum

' netplay.h: enum CONNECTION_STATUS
Public Enum CONNECTION_STATUS
    CONNECTIONSTATUS_PLAYER_DROPPED = 0
    CONNECTIONSTATUS_PLAYER_LEAVING = 1
    CONNECTIONSTATUS_DESYNC = 2
    CONNECTIONSTATUS_WAITING_FOR_PLAYER = 3

    CONNECTIONSTATUS_NORMAL = 4
End Enum

' netplay.h: enum LOBBY_ERROR_TYPES
Public Enum LOBBY_ERROR_TYPES
    ERROR_NOERROR
    ERROR_CONNECTION
    ERROR_FULL
    ERROR_INVALID
    ERROR_KICKED
    ERROR_WRONGVERSION
    ERROR_WRONGPASSWORD
    ERROR_HOSTDROPPED
    ERROR_WRONGDATA
    ERROR_UNKNOWNFILEISSUE
    ERROR_REDIRECT
End Enum

' netplay.h: enum NET_LOBBY_OPT_FIELD
Public Enum NET_LOBBY_OPT_FIELD
    NET_LOBBY_OPT_FIELD_INVALID
    NET_LOBBY_OPT_FIELD_GNAME
    NET_LOBBY_OPT_FIELD_MAPNAME
    NET_LOBBY_OPT_FIELD_HOSTNAME
    NET_LOBBY_OPT_FIELD_MAX
End Enum

' netplay.h: enum NetStatisticType
Public Enum NetStatisticType
    NetStatisticRawBytes
    NetStatisticUncompressedBytes
    NetStatisticPackets
End Enum

' netplay.h: enum class SpectatorToPlayerMoveResult
Public Enum SpectatorToPlayerMoveResult
    SpectatorToPlayerMoveResult_SUCCESS
    SpectatorToPlayerMoveResult_NEEDS_SLOT_SELECTION
    SpectatorToPlayerMoveResult_FAILED
End Enum

' netplay.h: enum class AsyncJoinApprovalAction
Public Enum AsyncJoinApprovalAction
    AsyncJoinApprovalAction_Approve
    AsyncJoinApprovalAction_ApproveSpectators
    AsyncJoinApprovalAction_Reject
End Enum

' netplay.h: enum class LobbyHostDirectJoinOption
Public Enum LobbyHostDirectJoinOption
    LobbyHostDirectJoinOption_None
    LobbyHostDirectJoinOption_Local
    LobbyHostDirectJoinOption_All
End Enum

' netpermissions.h: enum class ConnectPermissions
Public Enum ConnectPermissions
    ConnectPermissions_Blocked
    ConnectPermissions_Allowed
End Enum

' netplay.h: enum MESSAGE_TYPES
Public Enum MESSAGE_TYPES
    ' Net-related messages.
    NET_MIN_TYPE = 33               ' Minimum-1 valid NET_ type, *MUST* be first.
    NET_PING                        ' ping players.
    NET_PLAYER_STATS                ' player stats
    NET_TEXTMSG                     ' A simple text message between machines.
    NET_PLAYERRESPONDING            ' computer that sent this is now playing warzone!
    NET_OPTIONS                     ' welcome a player to a game.
    NET_KICK                        ' kick a player .
    NET_FIREUP                      ' campaign game has started, we can go too.. Shortcut message, not to be used in dmatch.
    NET_COLOURREQUEST               ' player requests a colour change.
    NET_FACTIONREQUEST              ' player requests a colour change.
    NET_AITEXTMSG                   ' chat between AIs
    NET_BEACONMSG                   ' place beacon
    NET_TEAMREQUEST                 ' request team membership
    NET_JOIN                        ' join a game
    NET_ACCEPTED                    ' accepted into game
    NET_PLAYER_INFO                 ' basic player info
    NET_PLAYER_JOINED               ' notice about player joining
    NET_PLAYER_LEAVING              ' A player is leaving, (nicely)
    NET_PLAYER_DROPPED              ' notice about player dropped / disconnected
    NET_GAME_FLAGS                  ' game flags
    NET_READY_REQUEST               ' player ready to start an mp game
    NET_REJECTED                    ' nope, you can't join
    NET_POSITIONREQUEST             ' position in GUI player list
    NET_DATA_CHECK                  ' Data integrity check
    NET_HOST_DROPPED                ' Host has dropped
    NET_SEND_TO_PLAYER              ' Non-host clients aren't directly connected to each other, so they talk via the host using these messages.
    NET_SHARE_GAME_QUEUE            ' Message contains a game message, which should be inserted into a queue.
    NET_FILE_REQUESTED              ' Player has requested a file (map/mod/?)
    NET_FILE_CANCELLED              ' Player cancelled a file request
    NET_FILE_PAYLOAD                ' sending file to the player that needs it
    NET_DEBUG_SYNC                  ' Synch error messages, so people don't have to use pastebin.
    NET_VOTE                        ' player vote
    NET_VOTE_REQUEST                ' Setup a vote popup
    NET_SPECTEXTMSG                 ' chat between spectators
    NET_PLAYERNAME_CHANGEREQUEST    ' non-host human player is changing their name.
    NET_PLAYER_SLOTTYPE_REQUEST     ' non-host human player is requesting a slot type change, or a host is asking a spectator if they want to play
    NET_PLAYER_SWAP_INDEX           ' a host-only message to move a player to another index
    NET_PLAYER_SWAP_INDEX_ACK       ' an acknowledgement message from a player whose index is being swapped
    NET_DATA_CHECK2                 ' Data2 integrity check
    NET_SECURED_NET_MESSAGE         ' A secured (+ authenticated) net message between two players
    NET_TEAM_STRATEGY               ' Player is sending an updated strategy notice to team members
    NET_QUICK_CHAT_MSG              ' Quick chat message
    NET_HOST_CONFIG                 ' Host configuration sent both before the game has started (in lobby), and after the game has started
    NET_MAX_TYPE                    ' Maximum+1 valid NET_ type, *MUST* be last.

    ' Game-state-related messages, must be processed by all clients at the same game time.
    GAME_MIN_TYPE = 111             ' Minimum-1 valid GAME_ type, *MUST* be first.
    GAME_DROIDINFO                  ' update a droid order.
    GAME_STRUCTUREINFO              ' Structure state.
    GAME_RESEARCHSTATUS             ' research state.
    GAME_TEMPLATE                   ' a new template
    GAME_TEMPLATEDEST               ' remove template
    GAME_ALLIANCE                   ' alliance data.
    GAME_GIFT                       ' a luvly gift between players.
    GAME_LASSAT                     ' lassat firing.
    GAME_GAME_TIME                  ' Game time. Used for synchronising, so that all messages are executed at the same gameTime on all clients.
    GAME_PLAYER_LEFT                ' Player has left or dropped.
    GAME_DROIDDISEMBARK             ' droid disembarked from a Transporter
    GAME_SYNC_REQUEST               ' Game event generated from scripts that is meant to be synced
    ' The following messages are used for debug mode.
    GAME_DEBUG_MODE                 ' Request enable/disable debug mode.
    GAME_DEBUG_ADD_DROID            ' Add droid.
    GAME_DEBUG_ADD_STRUCTURE        ' Add structure.
    GAME_DEBUG_ADD_FEATURE          ' Add feature.
    GAME_DEBUG_REMOVE_DROID         ' Remove droid.
    GAME_DEBUG_REMOVE_STRUCTURE     ' Remove structure.
    GAME_DEBUG_REMOVE_FEATURE       ' Remove feature.
    GAME_DEBUG_FINISH_RESEARCH      ' Research has been completed.
    ' End of debug messages.
    GAME_SYNC_OPT_CHANGE            ' Change synchronized options for a player (ex formation options)
    GAME_MAX_TYPE                   ' Maximum+1 valid GAME_ type, *MUST* be last.

    ' The following messages are used for playing back replays.
    REPLAY_ENDED = 255              ' A special message for signifying the end of the replay
    ' End of replay messages.
End Enum

' port_mapping_manager.h: enum class PortMappingDiscoveryStatus
Public Enum PortMappingDiscoveryStatus
    PortMappingDiscoveryStatus_NOT_STARTED = 0
    PortMappingDiscoveryStatus_IN_PROGRESS = 1
    PortMappingDiscoveryStatus_SUCCESS = 2
    PortMappingDiscoveryStatus_FAILURE = 3
    PortMappingDiscoveryStatus_TIMEOUT = 4
End Enum

' Name and version of this DLL, of the WzFramework it uses, and the game
' version reported by the exe through INetGame, for the log header.
Public Function NetBuildInfo() As String
    Dim gameVersion As String
    If Not netGame Is Nothing Then gameVersion = ", game " & netGame.versionGetVersionString()
    NetBuildInfo = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision & " (uses " & FrameworkBuildInfo() & gameVersion & ")"
End Function

' Hands over the exe's INetGame; Nothing at shutdown (the DLL holds a
' reference to an exe object, so the exe must release it before it ends).
Public Sub NetSetGame(ByVal game As INetGame)
    Set netGame = game
End Sub

'--------------------------------------------------------------------------
' netplay.h's NETPLAY NetPlay: its fields as NetPlay_<field>, a player's as NetPlay_players_<field>(i)

Public Property Get NetPlay_bComms() As Boolean
    netplay_c.ensureNetPlay
    NetPlay_bComms = netplay_c.NetPlay.bComms
End Property

Public Property Let NetPlay_bComms(ByVal value As Boolean)
    netplay_c.ensureNetPlay
    netplay_c.NetPlay.bComms = value
End Property

' NetPlay.playercount: uint32_t, the number of players in the game
Public Property Get NetPlay_playercount() As Long
    netplay_c.ensureNetPlay
    NetPlay_playercount = netplay_c.NetPlay.playercount
End Property

Public Property Let NetPlay_playercount(ByVal value As Long)
    netplay_c.ensureNetPlay
    netplay_c.NetPlay.playercount = value
End Property

Public Property Get NetPlay_isHost() As Boolean
    netplay_c.ensureNetPlay
    NetPlay_isHost = netplay_c.NetPlay.isHost
End Property

Public Property Let NetPlay_isHost(ByVal value As Boolean)
    netplay_c.ensureNetPlay
    netplay_c.NetPlay.isHost = value
End Property

Public Property Get NetPlay_hostPlayer() As Long
    netplay_c.ensureNetPlay
    NetPlay_hostPlayer = netplay_c.NetPlay.hostPlayer
End Property

Public Property Let NetPlay_hostPlayer(ByVal value As Long)
    netplay_c.ensureNetPlay
    netplay_c.NetPlay.hostPlayer = value
End Property

Public Property Get NetPlay_isHostAlive() As Boolean
    netplay_c.ensureNetPlay
    NetPlay_isHostAlive = netplay_c.NetPlay.isHostAlive
End Property

Public Property Let NetPlay_isHostAlive(ByVal value As Boolean)
    netplay_c.ensureNetPlay
    netplay_c.NetPlay.isHostAlive = value
End Property

Public Property Get NetPlay_isPortMappingEnabled() As Boolean
    netplay_c.ensureNetPlay
    NetPlay_isPortMappingEnabled = netplay_c.NetPlay.isPortMappingEnabled
End Property

Public Property Let NetPlay_isPortMappingEnabled(ByVal value As Boolean)
    netplay_c.ensureNetPlay
    netplay_c.NetPlay.isPortMappingEnabled = value
End Property

' NetPlay.gamePassword (char[password_string_size])
Public Function NetPlay_gamePassword() As String
    netplay_c.ensureNetPlay
    NetPlay_gamePassword = netplay_c.NetPlay.gamePassword
End Function

' NetPlay.GamePassworded
Public Function NetPlay_GamePassworded() As Boolean
    netplay_c.ensureNetPlay
    NetPlay_GamePassworded = netplay_c.NetPlay.GamePassworded
End Function

' NetPlay.players.size(): MAX_CONNECTED_PLAYERS, + 1 with a replay's spectator
Public Function NetPlay_players_size() As Long
    NetPlay_players_size = netplay_c.NetPlay_players_size()
End Function

' NetPlay.players.push_back(PLAYER())
Public Sub NetPlay_players_push_back()
    netplay_c.NetPlay_players_push_back
End Sub

' NetPlay.players.resize(n)
Public Sub NetPlay_players_resize(ByVal n As Long)
    netplay_c.NetPlay_players_resize n
End Sub

' NetPlay.players[i].allocated (Allocated as a human player)
Public Property Get NetPlay_players_allocated(ByVal i As Long) As Boolean
    NetPlay_players_allocated = netplay_c.NetPlay_players_allocated(i)
End Property

Public Property Let NetPlay_players_allocated(ByVal i As Long, ByVal value As Boolean)
    netplay_c.set_NetPlay_players_allocated i, value
End Property

' NetPlay.players[i].isSpectator (whether this slot is a spectator slot)
Public Property Get NetPlay_players_isSpectator(ByVal i As Long) As Boolean
    NetPlay_players_isSpectator = netplay_c.NetPlay_players_isSpectator(i)
End Property

Public Property Let NetPlay_players_isSpectator(ByVal i As Long, ByVal value As Boolean)
    netplay_c.set_NetPlay_players_isSpectator i, value
End Property

' NetPlay.players[i].faction (a FactionID)
Public Property Get NetPlay_players_faction(ByVal i As Long) As Long
    NetPlay_players_faction = netplay_c.NetPlay_players_faction(i)
End Property

Public Property Let NetPlay_players_faction(ByVal i As Long, ByVal value As Long)
    netplay_c.set_NetPlay_players_faction i, value
End Property

' NetPlay.players[i].position (int32_t: the map starting position)
Public Property Get NetPlay_players_position(ByVal i As Long) As Long
    NetPlay_players_position = netplay_c.NetPlay_players_position(i)
End Property

Public Property Let NetPlay_players_position(ByVal i As Long, ByVal value As Long)
    netplay_c.set_NetPlay_players_position i, value
End Property

' NetPlay.players[i].difficulty (an AIDifficulty: -2 HUMAN, -1 DISABLED, 0 SUPEREASY .. 4 INSANE)
Public Property Get NetPlay_players_difficulty(ByVal i As Long) As Long
    NetPlay_players_difficulty = netplay_c.NetPlay_players_difficulty(i)
End Property

Public Property Let NetPlay_players_difficulty(ByVal i As Long, ByVal value As Long)
    netplay_c.set_NetPlay_players_difficulty i, value
End Property

' NetPlay.players[i].name (char[StringSize])
Public Property Get NetPlay_players_name(ByVal i As Long) As String
    NetPlay_players_name = netplay_c.NetPlay_players_name(i)
End Property

Public Property Let NetPlay_players_name(ByVal i As Long, ByVal value As String)
    netplay_c.set_NetPlay_players_name i, value
End Property

' NetPlay.players[i].isAdmin (whether this slot has admin privs)
Public Property Get NetPlay_players_isAdmin(ByVal i As Long) As Boolean
    NetPlay_players_isAdmin = netplay_c.NetPlay_players_isAdmin(i)
End Property

Public Property Let NetPlay_players_isAdmin(ByVal i As Long, ByVal value As Boolean)
    netplay_c.set_NetPlay_players_isAdmin i, value
End Property

' NetPlay.players[i].colour (int32_t: Which colour slot this player is using)
Public Property Get NetPlay_players_colour(ByVal i As Long) As Long
    NetPlay_players_colour = netplay_c.NetPlay_players_colour(i)
End Property

Public Property Let NetPlay_players_colour(ByVal i As Long, ByVal value As Long)
    netplay_c.set_NetPlay_players_colour i, value
End Property

' NetPlay.players[i].heartattacktime (uint32_t: Time cardiac arrest started)
Public Property Get NetPlay_players_heartattacktime(ByVal i As Long) As Long
    NetPlay_players_heartattacktime = netplay_c.NetPlay_players_heartattacktime(i)
End Property

Public Property Let NetPlay_players_heartattacktime(ByVal i As Long, ByVal value As Long)
    netplay_c.set_NetPlay_players_heartattacktime i, value
End Property

' NetPlay.players[i].heartbeat (If connection is still alive or not)
Public Property Get NetPlay_players_heartbeat(ByVal i As Long) As Boolean
    NetPlay_players_heartbeat = netplay_c.NetPlay_players_heartbeat(i)
End Property

Public Property Let NetPlay_players_heartbeat(ByVal i As Long, ByVal value As Boolean)
    netplay_c.set_NetPlay_players_heartbeat i, value
End Property

' NetPlay.players[i].team (int32_t: Which team we are on (int32_t::max for spectator team))
Public Property Get NetPlay_players_team(ByVal i As Long) As Long
    NetPlay_players_team = netplay_c.NetPlay_players_team(i)
End Property

Public Property Let NetPlay_players_team(ByVal i As Long, ByVal value As Long)
    netplay_c.set_NetPlay_players_team i, value
End Property

' NetPlay.players[i].ready (player ready to start?)
Public Property Get NetPlay_players_ready(ByVal i As Long) As Boolean
    NetPlay_players_ready = netplay_c.NetPlay_players_ready(i)
End Property

Public Property Let NetPlay_players_ready(ByVal i As Long, ByVal value As Boolean)
    netplay_c.set_NetPlay_players_ready i, value
End Property

' NetPlay.players[i].ai (int8_t: index into sorted list of AIs, zero is always default AI)
Public Property Get NetPlay_players_ai(ByVal i As Long) As Long
    NetPlay_players_ai = netplay_c.NetPlay_players_ai(i)
End Property

Public Property Let NetPlay_players_ai(ByVal i As Long, ByVal value As Long)
    netplay_c.set_NetPlay_players_ai i, value
End Property

' NetPlay.players[i].autoGame (if we are running a autogame (AI controls us))
Public Property Get NetPlay_players_autoGame(ByVal i As Long) As Boolean
    NetPlay_players_autoGame = netplay_c.NetPlay_players_autoGame(i)
End Property

Public Property Let NetPlay_players_autoGame(ByVal i As Long, ByVal value As Boolean)
    netplay_c.set_NetPlay_players_autoGame i, value
End Property

' NetPlay.players[i].IPtextAddress (the joined player's IP address, as text)
Public Property Get NetPlay_players_IPtextAddress(ByVal i As Long) As String
    NetPlay_players_IPtextAddress = netplay_c.NetPlay_players_IPtextAddress(i)
End Property

Public Property Let NetPlay_players_IPtextAddress(ByVal i As Long, ByVal value As String)
    netplay_c.set_NetPlay_players_IPtextAddress i, value
End Property

' PLAYER::resetAll() of NetPlay.players[i]
Public Sub NetPlay_players_resetAll(ByVal i As Long)
    netplay_c.NetPlay_players_resetAll i
End Sub

' PLAYER::fileSendInProgress(): wzFiles not empty
Public Property Get NetPlay_players_fileSendInProgress(ByVal i As Long) As Boolean
    NetPlay_players_fileSendInProgress = netplay_c.NetPlay_players_fileSendInProgress(i)
End Property

' NetPlay.players[i].wzFiles (std::shared_ptr<std::vector<WZFile>>): its size, an item (0-based), push_back, clear;
' a hash as its 32 bytes
Public Function NetPlay_players_wzFiles_size(ByVal i As Long) As Long
    NetPlay_players_wzFiles_size = netplay_file_c.NetPlay_players_wzFiles_size(i)
End Function

Public Function NetPlay_players_wzFiles_item(ByVal i As Long, ByVal n As Long) As WZFile
    Set NetPlay_players_wzFiles_item = netplay_file_c.NetPlay_players_wzFiles_item(i, n)
End Function

Public Sub NetPlay_players_wzFiles_push_back(ByVal i As Long, ByVal file As WZFile)
    netplay_file_c.NetPlay_players_wzFiles_push_back i, file
End Sub

' wzFiles->emplace_back(pFileHandle, filename, hash, fileSize)
Public Sub NetPlay_players_wzFiles_emplace_back(ByVal i As Long, ByVal pFileHandle As Long, ByVal filename As String, hashBytes() As Byte, ByVal fileSize As Long)
    netplay_file_c.NetPlay_players_wzFiles_emplace_back i, pFileHandle, filename, hashBytes, fileSize
End Sub

Public Sub NetPlay_players_wzFiles_clear(ByVal i As Long)
    netplay_file_c.NetPlay_players_wzFiles_clear i
End Sub

' any_of(file.hash == hash)
Public Function NetPlay_players_wzFiles_hasHash(ByVal i As Long, hashBytes() As Byte) As Boolean
    NetPlay_players_wzFiles_hasHash = netplay_file_c.NetPlay_players_wzFiles_hasHash(i, hashBytes)
End Function

' erase the files with this hash (NET_FILE_CANCELLED)
Public Sub NetPlay_players_wzFiles_eraseHash(ByVal i As Long, hashBytes() As Byte)
    netplay_file_c.NetPlay_players_wzFiles_eraseHash i, hashBytes
End Sub

' erase the files whose handle() is nullptr (sent)
Public Sub NetPlay_players_wzFiles_eraseClosed(ByVal i As Long)
    netplay_file_c.NetPlay_players_wzFiles_eraseClosed i
End Sub

' NetPlay.scriptSetPlayerDataStrings[idx] = obj.get<std::unordered_map<std::string, std::string>>()
Public Sub NetPlay_scriptSetPlayerDataStrings_assign(ByVal idx As Long, ByVal obj As nlohmann_json)
    netplay_c.NetPlay_scriptSetPlayerDataStrings_assign idx, obj
End Sub

' NetPlay.scriptSetPlayerDataStrings.size()
Public Function NetPlay_scriptSetPlayerDataStrings_size() As Long
    NetPlay_scriptSetPlayerDataStrings_size = netplay_c.NetPlay_scriptSetPlayerDataStrings_size()
End Function

' NetPlay.scriptSetPlayerDataStrings[i].clear()
Public Sub NetPlay_scriptSetPlayerDataStrings_clear(ByVal i As Long)
    netplay_c.NetPlay_scriptSetPlayerDataStrings_clear i
End Sub

' NetPlay.scriptSetPlayerDataStrings[i]: its size and its entries (n: 0 .. size - 1, in insertion order)
Public Function NetPlay_scriptSetPlayerDataStrings_count(ByVal i As Long) As Long
    NetPlay_scriptSetPlayerDataStrings_count = netplay_c.NetPlay_scriptSetPlayerDataStrings_count(i)
End Function

Public Sub NetPlay_scriptSetPlayerDataStrings_entry(ByVal i As Long, ByVal n As Long, key As String, value As String)
    netplay_c.NetPlay_scriptSetPlayerDataStrings_entry i, n, key, value
End Sub

' NetPlay.playerReferences[i] (std::shared_ptr<PlayerReference>)
Public Function NetPlay_playerReferences(ByVal i As Long) As PlayerReference
    Set NetPlay_playerReferences = netplay_c.NetPlay_playerReferences(i)
End Function

' NetPlay.playerReferences[i]->disconnect(); NetPlay.playerReferences[i] = std::make_shared<PlayerReference>(i)
Public Sub NetPlay_playerReferences_reset(ByVal i As Long)
    netplay_c.NetPlay_playerReferences_reset i
End Sub

'--------------------------------------------------------------------------
' netplay.cpp's settings, netlobby.h

Public Property Get netGameserverPortOverride() As Boolean
    netGameserverPortOverride = netplay_c.netGameserverPortOverride
End Property

Public Property Let netGameserverPortOverride(ByVal value As Boolean)
    netplay_c.netGameserverPortOverride = value
End Property

Public Property Get iptoconnect() As String
    iptoconnect = netplay_c.iptoconnect
End Property

Public Property Let iptoconnect(ByVal value As String)
    netplay_c.iptoconnect = value
End Property

Public Property Get cliConnectAsSpectator() As Boolean
    cliConnectAsSpectator = netplay_c.cliConnectAsSpectator
End Property

Public Property Let cliConnectAsSpectator(ByVal value As Boolean)
    netplay_c.cliConnectAsSpectator = value
End Property

Public Function netlobby_GetDefaultLobbyAddress() As String
    netlobby_GetDefaultLobbyAddress = netlobby_c.netlobby_GetDefaultLobbyAddress()
End Function

Public Sub NETsetLobbyserverAddress(ByVal lobbyAddress As String)
    netplay_c.NETsetLobbyserverAddress lobbyAddress
End Sub

Public Function NETgetLobbyserverAddress() As String
    NETgetLobbyserverAddress = netplay_c.NETgetLobbyserverAddress()
End Function

Public Sub NETsetGameserverPort(ByVal port As Long)
    netplay_c.NETsetGameserverPort port
End Sub

Public Function NETgetGameserverPort() As Long
    NETgetGameserverPort = netplay_c.NETgetGameserverPort()
End Function

Public Sub NETsetJoinPreferenceIPv6(ByVal bTryIPv6First As Boolean)
    netplay_c.NETsetJoinPreferenceIPv6 bTryIPv6First
End Sub

Public Function NETgetJoinPreferenceIPv6() As Boolean
    NETgetJoinPreferenceIPv6 = netplay_c.NETgetJoinPreferenceIPv6()
End Function

Public Sub NETsetDefaultMPHostFreeChatPreference(ByVal enabled As Boolean)
    netplay_c.NETsetDefaultMPHostFreeChatPreference enabled
End Sub

Public Function NETgetDefaultMPHostFreeChatPreference() As Boolean
    NETgetDefaultMPHostFreeChatPreference = netplay_c.NETgetDefaultMPHostFreeChatPreference()
End Function

Public Sub NETsetEnableTCPNoDelay(ByVal enabled As Boolean)
    netplay_c.NETsetEnableTCPNoDelay enabled
End Sub

Public Function NETgetEnableTCPNoDelay() As Boolean
    NETgetEnableTCPNoDelay = netplay_c.NETgetEnableTCPNoDelay()
End Function

' netlobby::HostJoinOptions NETgetLobbyHostJoinOptions(): its three fields (HostJoinOptionValues)
Public Sub NETgetLobbyHostJoinOptions(botProt As Long, proxyIPs As Long, hostingIPs As Long)
    netplay_c.NETgetLobbyHostJoinOptions botProt, proxyIPs, hostingIPs
End Sub

Public Sub NETsetLobbyHostJoinOptions(ByVal botProt As Long, ByVal proxyIPs As Long, ByVal hostingIPs As Long)
    netplay_c.NETsetLobbyHostJoinOptions botProt, proxyIPs, hostingIPs
End Sub

Public Function NETgetLobbyHostDirectJoinOption() As LobbyHostDirectJoinOption
    NETgetLobbyHostDirectJoinOption = netplay_c.NETgetLobbyHostDirectJoinOption()
End Function

Public Sub NETsetLobbyHostDirectJoinOption(ByVal opt As LobbyHostDirectJoinOption)
    netplay_c.NETsetLobbyHostDirectJoinOption opt
End Sub

' optional<bool> NET_getHostJoinOptionSupportedByLobby(netlobby::HostJoinOptionType optType): has_value, the value ByRef
Public Function NET_getHostJoinOptionSupportedByLobby(ByVal optType As HostJoinOptionType, result As Boolean) As Boolean
    NET_getHostJoinOptionSupportedByLobby = netplay_c.NET_getHostJoinOptionSupportedByLobby(optType, result)
End Function

Public Function NET_getCurrentHostedLobbyGameId() As String
    NET_getCurrentHostedLobbyGameId = netplay_c.NET_getCurrentHostedLobbyGameId()
End Function

Public Function NET_getLobbyDisabled() As Boolean
    NET_getLobbyDisabled = netplay_c.NET_getLobbyDisabled()
End Function

Public Function NET_getLobbyDisabledInfoLinkURL() As String
    NET_getLobbyDisabledInfoLinkURL = netplay_c.NET_getLobbyDisabledInfoLinkURL()
End Function

Public Sub NET_setLobbyDisabled(ByVal infoLinkURL As String)
    netplay_c.NET_setLobbyDisabled infoLinkURL
End Sub

Public Sub NETsetLobbyOptField(ByVal value As String, ByVal field As NET_LOBBY_OPT_FIELD)
    netplay_c.NETsetLobbyOptField value, field
End Sub

Public Sub NETsetLobbyConfigFlagsFields(ByVal alliancesType As Long, ByVal techLevel As Long, ByVal powerLevel As Long, ByVal basesLevel As Long)
    netplay_conn_c.NETsetLobbyConfigFlagsFields alliancesType, techLevel, powerLevel, basesLevel
End Sub

' ingameLimitFlags: ingame.flags (uint8_t)
Public Sub NETsetLobbyLimitFlags(ByVal ingameLimitFlags As Long)
    netplay_conn_c.NETsetLobbyLimitFlags ingameLimitFlags
End Sub

' SpectatorInfo SpectatorInfo::currentNetPlayState(), packed as its toUint32(): spectatorsJoined in the high 16
' bits, totalSpectatorSlots in the low 16
Public Function SpectatorInfo_currentNetPlayState() As Long
    SpectatorInfo_currentNetPlayState = netplay_c.SpectatorInfo_currentNetPlayState()
End Function

' SpectatorInfo NETGameGetSpectatorInfo() (packed, toUint32)
Public Function NETGameGetSpectatorInfo() As Long
    NETGameGetSpectatorInfo = netplay_c.NETGameGetSpectatorInfo()
End Function

'--------------------------------------------------------------------------
' sync_debug.h

' syncDebug(...): the macro, which passes __FUNCTION__ to _syncDebug
Public Sub syncDebug(ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    sync_debug_c.syncDebugV function_, str_, a
End Sub

' _syncDebugIntList(function, str, ints, numInts): faster than syncDebug; str takes ints only
Public Sub syncDebugIntList_(ByVal function_ As String, ByVal str_ As String, ints() As Long, ByVal numInts As Long)
    sync_debug_c.syncDebugIntList_ function_, str_, ints, numInts
End Sub

' _syncDebugBacktrace(function): the syncDebugBacktrace() macro passes __FUNCTION__
Public Sub syncDebugBacktrace_(ByVal function_ As String)
    sync_debug_c.syncDebugBacktrace_ function_
End Sub

Public Function syncDebugGetCrc() As Long
    syncDebugGetCrc = sync_debug_c.syncDebugGetCrc()
End Function

Public Sub syncDebugSetCrc(ByVal crc As Long)
    sync_debug_c.syncDebugSetCrc crc
End Sub

Public Sub resetSyncDebug()
    sync_debug_c.resetSyncDebug
End Sub

' GameCrcType nextDebugSync(): gameTime_ is gtime's gameTime
Public Function nextDebugSync(ByVal gameTime_ As Long) As Long
    nextDebugSync = sync_debug_c.nextDebugSync(gameTime_)
End Function

Public Sub NET_setDebuggingModeVerboseOutputAllSyncLogs(ByVal untilGameTime As Long)
    sync_debug_c.NET_setDebuggingModeVerboseOutputAllSyncLogs untilGameTime
End Sub

' debugVerboseLogSyncIfNeeded(): gameTime_ is gtime's gameTime
Public Sub debugVerboseLogSyncIfNeeded(ByVal gameTime_ As Long)
    sync_debug_c.debugVerboseLogSyncIfNeeded gameTime_
End Sub

' (VB, for the tests) the static dumpLocalDebugSyncLog of the log being written
Public Function dumpCurrentDebugSyncLog() As String
    dumpCurrentDebugSyncLog = sync_debug_c.dumpCurrentDebugSyncLog()
End Function

Public Function checkDebugSync(ByVal checkGameTime As Long, ByVal checkCrc As Long) As Boolean
    checkDebugSync = sync_debug_c.checkDebugSync(checkGameTime, checkCrc)
End Function

' (VB, for the tests) the static dumpLocalDebugSyncLog(logIndex), and syncDebugNext
Public Function dumpLocalDebugSyncLog(ByVal logIndex As Long) As String
    dumpLocalDebugSyncLog = sync_debug_c.dumpLocalDebugSyncLog(logIndex)
End Function

Public Function syncDebugNextIndex() As Long
    syncDebugNextIndex = sync_debug_c.syncDebugNextIndex()
End Function

'--------------------------------------------------------------------------
' netplay.h: start and end, the statistics, the players

' int NETinit(ConnectionProviderType pt)
Public Function NETinit(ByVal pt As Long) As Long
    NETinit = netplay_c.NETinit(pt)
End Function

Public Function NETshutdown() As Long
    NETshutdown = netplay_c.NETshutdown()
End Function

Public Function netplayShutDown() As Boolean
    netplayShutDown = netplay_c.netplayShutDown()
End Function

' int NETclose(): close the socket
Public Function NETclose() As Long
    NETclose = netplay_conn_c.NETclose()
End Function

' size_t NETgetStatistic(NetStatisticType type, bool sent, bool isTotal = false)
Public Function NETgetStatistic(ByVal typ As Long, ByVal sent As Boolean, Optional ByVal isTotal As Boolean = False) As Double
    NETgetStatistic = netplay_c.NETgetStatistic(typ, sent, isTotal)
End Function

Public Sub NETaddRedirects()
    netplay_c.NETaddRedirects
End Sub

Public Sub NETremRedirects()
    netplay_c.NETremRedirects
End Sub

Public Sub NETinitPortMapping()
    netplay_c.NETinitPortMapping
End Sub

' multiint.cpp's `if (!ipv4MappingRequest)`: whether there is a port mapping request
Public Function ipv4MappingRequest_exists() As Boolean
    ipv4MappingRequest_exists = Not netplay_c.ipv4MappingRequest Is Nothing
End Function

' PortMappingManager::instance().get_status(ipv4MappingRequest)
Public Function PortMappingManager_get_status_ipv4MappingRequest() As PortMappingDiscoveryStatus
    PortMappingManager_get_status_ipv4MappingRequest = port_mapping_manager_c.PortMappingManager_get_status(netplay_c.ipv4MappingRequest)
End Function

Public Sub NET_InitPlayer(ByVal i As Long, ByVal initPosition As Boolean, Optional ByVal initTeams As Boolean = False, Optional ByVal initSpectator As Boolean = False)
    netplay_c.NET_InitPlayer i, initPosition, initTeams, initSpectator
End Sub

Public Sub NET_InitPlayers(Optional ByVal initTeams As Boolean = False, Optional ByVal initSpectator As Boolean = False)
    netplay_c.NET_InitPlayers initTeams, initSpectator
End Sub

' uint8_t NET_numHumanPlayers(void)
Public Function NET_numHumanPlayers() As Long
    NET_numHumanPlayers = netplay_c.NET_numHumanPlayers()
End Function

' std::vector<uint8_t> NET_getHumanPlayers(void): RetVal(0 .. return value - 1)
Public Function NET_getHumanPlayers(RetVal() As Long) As Long
    NET_getHumanPlayers = netplay_c.NET_getHumanPlayers(RetVal)
End Function

Public Sub NETplayerKicked(ByVal index As Long, Optional ByVal quiet As Boolean = False)
    netplay_conn_c.NETplayerKicked index, quiet
End Sub

Public Function NETplayerHasConnection(ByVal index As Long) As Boolean
    NETplayerHasConnection = netplay_c.NETplayerHasConnection(index)
End Function

' bool NETchangePlayerName(UDWORD player, char *newName)
Public Function NETchangePlayerName(ByVal index As Long, ByVal newName As String) As Boolean
    NETchangePlayerName = netplay_c.NETchangePlayerName(index, newName)
End Function

Public Sub NETfixDuplicatePlayerNames()
    netplay_c.NETfixDuplicatePlayerNames
End Sub

Public Function NETgetGameFlags(ByVal flag As Long) As Long
    NETgetGameFlags = netplay_c.NETgetGameFlags(flag)
End Function

Public Sub NETsetGameFlags(ByVal flag As Long, ByVal value As Long)
    netplay_c.NETsetGameFlags flag, value
End Sub

Public Function NETcanOpenNewSpectatorSlot() As Boolean
    NETcanOpenNewSpectatorSlot = netplay_sys_c.NETcanOpenNewSpectatorSlot()
End Function

Public Function NETopenNewSpectatorSlot() As Boolean
    NETopenNewSpectatorSlot = netplay_sys_c.NETopenNewSpectatorSlot()
End Function

Public Function NETmovePlayerToSpectatorOnlySlot(ByVal playerIdx As Long, Optional ByVal hostOverride As Boolean = False) As Boolean
    NETmovePlayerToSpectatorOnlySlot = netplay_sys_c.NETmovePlayerToSpectatorOnlySlot(playerIdx, hostOverride)
End Function

' SpectatorToPlayerMoveResult NETmoveSpectatorToPlayerSlot(uint32_t playerIdx, optional<uint32_t> newPlayerIdx,
' bool hostOverride = false): newPlayerIdx -1 = nullopt
Public Function NETmoveSpectatorToPlayerSlot(ByVal playerIdx As Long, ByVal newPlayerIdx As Long, Optional ByVal hostOverride As Boolean = False) As SpectatorToPlayerMoveResult
    NETmoveSpectatorToPlayerSlot = netplay_sys_c.NETmoveSpectatorToPlayerSlot(playerIdx, newPlayerIdx <> -1, newPlayerIdx, hostOverride)
End Function

' Sets if the game is password protected or not
Public Sub NETGameLocked(ByVal flag As Boolean)
    netplay_c.NETGameLocked flag
End Sub

Public Function NETGameIsLocked() As Boolean
    NETGameIsLocked = netplay_c.NETGameIsLocked()
End Function

' Sets the game password
Public Sub NETsetGamePassword(ByVal password As String)
    netplay_c.NETsetGamePassword password
End Sub

Public Sub NETresetGamePassword()
    netplay_c.NETresetGamePassword
End Sub

Public Sub NETsetAsyncJoinApprovalRequired(ByVal enabled As Boolean)
    netplay_c.NETsetAsyncJoinApprovalRequired enabled
End Sub

Public Function NETgetAsyncJoinApprovalRequired() As Boolean
    NETgetAsyncJoinApprovalRequired = netplay_c.NETgetAsyncJoinApprovalRequired()
End Function

' bool NETsetAsyncJoinApprovalResult(const std::string& uniqueJoinID, AsyncJoinApprovalAction action,
' optional<uint8_t> explicitPlayerIdx = nullopt, LOBBY_ERROR_TYPES rejectedReason = ERROR_NOERROR,
' optional<std::string> customRejectionMessage = nullopt)
Public Function NETsetAsyncJoinApprovalResult(ByVal uniqueJoinID As String, ByVal action As AsyncJoinApprovalAction, ByVal hasExplicitPlayerIdx As Boolean, _
        ByVal explicitPlayerIdx As Long, Optional ByVal rejectedReason As LOBBY_ERROR_TYPES = ERROR_NOERROR, _
        Optional ByVal hasCustomRejectionMessage As Boolean = False, Optional ByVal customRejectionMessage As String = "") As Boolean
    NETsetAsyncJoinApprovalResult = netplay_c.NETsetAsyncJoinApprovalResult(uniqueJoinID, action, hasExplicitPlayerIdx, explicitPlayerIdx, _
        rejectedReason, hasCustomRejectionMessage, customRejectionMessage)
End Function

' the netcode version a replay or a joining client must have
Public Function NETisCorrectVersion(ByVal game_version_major As Long, ByVal game_version_minor As Long) As Boolean
    NETisCorrectVersion = netplay_c.NETisCorrectVersion(game_version_major, game_version_minor)
End Function

Public Function NETisGreaterVersion(ByVal game_version_major As Long, ByVal game_version_minor As Long) As Boolean
    NETisGreaterVersion = netplay_c.NETisGreaterVersion(game_version_major, game_version_minor)
End Function

Public Function NETGetMajorVersion() As Long
    NETGetMajorVersion = netplay_c.NETGetMajorVersion()
End Function

Public Function NETGetMinorVersion() As Long
    NETGetMinorVersion = netplay_c.NETGetMinorVersion()
End Function

'--------------------------------------------------------------------------
' netplay.h: the message traffic

Public Function NETrecvNet(queue As Long, typ As Long) As Boolean
    NETrecvNet = netplay_conn_c.NETrecvNet(queue, typ)
End Function

Public Function NETrecvGame(queue As Long, typ As Long) As Boolean
    NETrecvGame = netplay_conn_c.NETrecvGame(queue, typ)
End Function

' (VB) and the port mapping's pass (libplum's client thread, PortMappingManager's monitoring thread: no threads):
' the title loop and the game loop call NETflush each frame
Public Sub NETflush()
    netplay_conn_c.NETflush
    port_mapping_manager_c.PortMappingManager_poll
End Sub

' bool NETsend(NETQUEUE queue, NetMessage const &message)
Public Function NETsend(ByVal queue As Long, ByVal message As NetMessage) As Boolean
    NETsend = netplay_conn_c.NETsend(queue, message)
End Function

Public Function messageTypeToString(ByVal messageType_ As Long) As String
    messageTypeToString = netplay_c.messageTypeToString(messageType_)
End Function

Public Sub NETsetPlayerConnectionStatus(ByVal status As Long, ByVal player As Long)
    netplay_c.NETsetPlayerConnectionStatus status, player
End Sub

Public Function NETcheckPlayerConnectionStatus(ByVal status As Long, ByVal player As Long) As Boolean
    NETcheckPlayerConnectionStatus = netplay_c.NETcheckPlayerConnectionStatus(status, player)
End Function

' bool netPlayersUpdated (update flags)
Public Property Get netPlayersUpdated() As Boolean
    netPlayersUpdated = netplay_c.netPlayersUpdated
End Property

Public Property Let netPlayersUpdated(ByVal value As Boolean)
    netplay_c.netPlayersUpdated = value
End Property

Public Sub NETBroadcastPlayerInfo(ByVal index As Long)
    netplay_c.NETBroadcastPlayerInfo index
End Sub

Public Sub NETBroadcastTwoPlayerInfo(ByVal index1 As Long, ByVal index2 As Long)
    netplay_c.NETBroadcastTwoPlayerInfo index1, index2
End Sub

Public Sub NETSendAllPlayerInfoTo(ByVal to_ As Long)
    netplay_c.NETSendAllPlayerInfoTo to_
End Sub

'--------------------------------------------------------------------------
' netplay.h: hosting and joining

' bool NEThostGame(const char *SessionName, const char *PlayerName, const EcKey& playerIdentity, bool spectatorHost,
' uint32_t gameType, uint32_t blindMode, UDWORD plyrs, uint16_t defaultOpenSpectatorSlots, uint8_t alliancesType,
' uint8_t techLevel, uint8_t powerLevel, uint8_t basesLevel), without the playerIdentity (the exe's signature):
' WzNet takes getMultiStats(selectedPlayer).identity through INetGame, the identity multiopt.cpp passes
Public Function NEThostGame(ByVal SessionName As String, ByVal PlayerName As String, ByVal spectatorHost As Boolean, _
                            ByVal gameType As Long, ByVal blindMode As Long, ByVal plyrs As Long, ByVal defaultOpenSpectatorSlots As Long, _
                            ByVal alliancesType As Long, ByVal techLevel As Long, ByVal powerLevel As Long, ByVal basesLevel As Long) As Boolean
    Dim key() As Byte, n As Long
    n = netGame.getMultiStatsIdentity(selectedPlayer, key)
    NEThostGame = netplay_conn_c.NEThostGame(SessionName, PlayerName, key, n, spectatorHost, gameType, blindMode, plyrs, defaultOpenSpectatorSlots, _
                                             alliancesType, techLevel, powerLevel, basesLevel)
End Function

' Stop the dplay interface from accepting more players.
Public Function NEThaltJoining(ByVal matchStarted As Boolean) As Boolean
    NEThaltJoining = netplay_conn_c.NEThaltJoining(matchStarted)
End Function

' bool NETpromoteJoinAttemptToEstablishedConnectionToHost(uint32_t hostPlayer, uint8_t index, const char* playerName,
' NETQUEUE joiningQUEUE, IClientConnection** client_joining_socket, IConnectionPollGroup** client_joining_socket_set):
' consumes the socket and the socket set (sets them to Nothing)
Public Function NETpromoteJoinAttemptToEstablishedConnectionToHost(ByVal hostPlayer As Long, ByVal index As Long, ByVal playerName As String, _
        ByVal joiningQUEUE As Long, client_joining_socket As IClientConnection, client_joining_socket_set As IConnectionPollGroup) As Boolean
    NETpromoteJoinAttemptToEstablishedConnectionToHost = netplay_conn_c.NETpromoteJoinAttemptToEstablishedConnectionToHost(hostPlayer, index, playerName, _
        joiningQUEUE, client_joining_socket, client_joining_socket_set)
End Function

' uint32_t NETgetJoinConnectionNETPINGChallengeFromHostSize() / ...FromClientSize()
Public Function NETgetJoinConnectionNETPINGChallengeFromHostSize() As Long
    NETgetJoinConnectionNETPINGChallengeFromHostSize = netplay_c.NETgetJoinConnectionNETPINGChallengeFromHostSize()
End Function

Public Function NETgetJoinConnectionNETPINGChallengeFromClientSize() As Long
    NETgetJoinConnectionNETPINGChallengeFromClientSize = netplay_c.NETgetJoinConnectionNETPINGChallengeFromClientSize()
End Function

Public Sub NETacceptIncomingConnections()
    netplay_conn_c.NETacceptIncomingConnections
End Sub

Public Sub NETadjustConnectedTimeoutForClients()
    netplay_conn_c.NETadjustConnectedTimeoutForClients
End Sub

' optional<std::string> NET_getCurrentHostTextAddress(): has_value, the value ByRef
Public Function NET_getCurrentHostTextAddress(result As String) As Boolean
    NET_getCurrentHostTextAddress = netplay_conn_c.NET_getCurrentHostTextAddress(result)
End Function

'--------------------------------------------------------------------------
' netplay.h: the file transfers; wzfile.h

' int NETsendFile(WZFile &file, unsigned player): % of the file sent
Public Function NETsendFile(ByVal file As WZFile, ByVal player As Long) As Long
    NETsendFile = netplay_file_c.NETsendFile(file, player)
End Function

' NETsendFile((*NetPlay.players[player].wzFiles)[fileIdx], player)
Public Function NETsendFile_idx(ByVal player As Long, ByVal fileIdx As Long) As Long
    NETsendFile_idx = netplay_file_c.NETsendFile_idx(player, fileIdx)
End Function

' int NETrecvFile(NETQUEUE queue): % of the file received so far
Public Function NETrecvFile(ByVal queue As Long) As Long
    NETrecvFile = netplay_file_c.NETrecvFile(queue)
End Function

' unsigned NETgetDownloadProgress(unsigned player): Returns 100 when done.
Public Function NETgetDownloadProgress(ByVal player As Long) As Long
    NETgetDownloadProgress = netplay_file_c.NETgetDownloadProgress(player)
End Function

' const std::vector<WZFile>& NET_getDownloadingWzFiles(): its size, its items (0-based); empty()
Public Function NET_getDownloadingWzFiles_size() As Long
    NET_getDownloadingWzFiles_size = netplay_file_c.NET_getDownloadingWzFiles_size()
End Function

Public Function NET_getDownloadingWzFiles_item(ByVal i As Long) As WZFile
    Set NET_getDownloadingWzFiles_item = netplay_file_c.NET_getDownloadingWzFiles_item(i)
End Function

Public Function NET_getDownloadingWzFiles_empty() As Boolean
    NET_getDownloadingWzFiles_empty = (netplay_file_c.NET_getDownloadingWzFiles_size() = 0)
End Function

' void NET_addDownloadingWZFile(WZFile&& newFile), as multiopt.cpp calls it: NET_addDownloadingWZFile(WZFile(pFileHandle,
' filename, hash)), the hash's 32 bytes
Public Sub NET_addDownloadingWZFile(ByVal pFileHandle As Long, ByVal filename As String, hashBytes() As Byte, Optional ByVal size As Long = 0)
    netplay_file_c.NET_addDownloadingWZFile wzfile_c.WZFile_new(pFileHandle, filename, hashBytes, size)
End Sub

' std::any_of(NET_getDownloadingWzFiles().begin(), ...end(), [&hash](WZFile const &file) { return file.hash == hash; })
Public Function NET_getDownloadingWzFiles_hasHash(hashBytes() As Byte) As Boolean
    NET_getDownloadingWzFiles_hasHash = netplay_file_c.NET_getDownloadingWzFiles_hasHash(hashBytes)
End Function

Public Sub NET_clearDownloadingWZFiles()
    netplay_file_c.NET_clearDownloadingWZFiles
End Sub

' WZFile(PHYSFS_file* handle, const std::string& filename, Sha256 hash, uint32_t size = 0): the hash's 32 bytes
Public Function WZFile_new(ByVal handle As Long, ByVal filename As String, hashBytes() As Byte, Optional ByVal size As Long = 0) As WZFile
    Set WZFile_new = wzfile_c.WZFile_new(handle, filename, hashBytes, size)
End Function

'--------------------------------------------------------------------------
' netlog.h, netpermissions.h

Public Function NETlogEntry(ByVal str_ As String, ByVal a As Long, ByVal b As Long) As Boolean
    NETlogEntry = netlog_c.NETlogEntry(str_, a, b)
End Function

Public Sub addIPToBanList(ByVal ip As String, ByVal name_ As String)
    netpermissions_c.addIPToBanList ip, name_
End Sub

Public Function removeIPFromBanList(ByVal ip As String) As Boolean
    removeIPFromBanList = netpermissions_c.removeIPFromBanList(ip)
End Function

Public Function onBanList(ByVal ip As String) As Boolean
    onBanList = netpermissions_c.onBanList(ip)
End Function

' std::vector<PLAYER_IP> NETgetIPBanList(): its size, an entry's pname and IPAddress
Public Function NETgetIPBanList_size() As Long
    NETgetIPBanList_size = netpermissions_c.NETgetIPBanList_size()
End Function

Public Sub NETgetIPBanList_item(ByVal k As Long, pname As String, IPAddress As String)
    netpermissions_c.NETgetIPBanList_item k, pname, IPAddress
End Sub

' identityStr: Either the the (b64) public key (preferred) or the public hash
Public Sub netPermissionsSet_Connect(ByVal identityStr As String, ByVal perm As ConnectPermissions)
    netpermissions_c.netPermissionsSet_Connect identityStr, perm
End Sub

Public Function netPermissionsUnset_Connect(ByVal identityStr As String) As Boolean
    netPermissionsUnset_Connect = netpermissions_c.netPermissionsUnset_Connect(identityStr)
End Function

Public Function netPermissionsRemoveAll(ByVal identityStr As String) As Boolean
    netPermissionsRemoveAll = netpermissions_c.netPermissionsRemoveAll(identityStr)
End Function

' optional<ConnectPermissions> netPermissionsCheck_Connect(const std::string& identityStr): has_value, the value ByRef
Public Function netPermissionsCheck_Connect(ByVal identityStr As String, perm As Long) As Boolean
    netPermissionsCheck_Connect = netpermissions_c.netPermissionsCheck_Connect(identityStr, perm)
End Function

'--------------------------------------------------------------------------
' netjoin.h

Public Function JoinConnectionDescription_connectiontype_to_string(ByVal typ As Long) As String
    JoinConnectionDescription_connectiontype_to_string = netjoin_c.JoinConnectionDescription_connectiontype_to_string(typ)
End Function

' optional<JoinConnectionType> connectiontype_from_string(const std::string& str): has_value, the value ByRef
Public Function JoinConnectionDescription_connectiontype_from_string(ByVal str_ As String, typ As Long) As Boolean
    JoinConnectionDescription_connectiontype_from_string = netjoin_c.JoinConnectionDescription_connectiontype_from_string(str_, typ)
End Function

'--------------------------------------------------------------------------
' connection_provider_registry.h, error_categories.h, pending_writes_manager_map.h (the connection layer)

' std::shared_ptr<WzConnectionProvider> ConnectionProviderRegistry::Instance().Get(pt)
Public Function ConnectionProviderRegistry_Get(ByVal pt As Long) As WzConnectionProvider
    Set ConnectionProviderRegistry_Get = connection_provider_registry_c.ConnectionProviderRegistry_Get(pt)
End Function

Public Function ConnectionProviderRegistry_IsRegistered(ByVal pt As Long) As Boolean
    ConnectionProviderRegistry_IsRegistered = connection_provider_registry_c.ConnectionProviderRegistry_IsRegistered(pt)
End Function

Public Sub ConnectionProviderRegistry_Register(ByVal pt As Long)
    connection_provider_registry_c.ConnectionProviderRegistry_Register pt
End Sub

Public Sub ConnectionProviderRegistry_Shutdown()
    connection_provider_registry_c.ConnectionProviderRegistry_Shutdown
End Sub

' PendingWritesManager::pollAll (the writer thread's pass; netplay runs it at NETrecvNet, NETrecvGame, NETflush)
Public Sub PendingWritesManager_pollAll()
    pending_writes_manager_map_c.PendingWritesManager_pollAll
End Sub

' std::error_code (an ec Long: the category in the high byte, the value in the low 24 bits): message(), value(),
' the category, ec == std::errc::x
Public Function net_errorMessage(ByVal ec As Long) As String
    net_errorMessage = error_categories_c.net_errorMessage(ec)
End Function

Public Function net_errorValue(ByVal ec As Long) As Long
    net_errorValue = error_categories_c.net_errorValue(ec)
End Function

Public Function net_errorCategory(ByVal ec As Long) As Long
    net_errorCategory = error_categories_c.net_errorCategory(ec)
End Function

Public Function net_errorEqualsErrc(ByVal ec As Long, ByVal errcValue As Long) As Boolean
    net_errorEqualsErrc = error_categories_c.net_errorEqualsErrc(ec, errcValue)
End Function

' std::error_code make_network_error_code(int errnoValue)
Public Function make_network_error_code(ByVal errnoValue As Long) As Long
    make_network_error_code = error_categories_c.make_network_error_code(errnoValue)
End Function

'--------------------------------------------------------------------------
' nettypes.h, netqueue.h: the queues. A NETQUEUE is a Long:
' queueType * &H10000 + index * &H100 + exclude.

Public Function NETQUEUE_make(ByVal queueType_ As Long, ByVal index As Long, ByVal exclude As Long) As Long
    NETQUEUE_make = nettypes_c.NETQUEUE_make(queueType_, index, exclude)
End Function

Public Function NETQUEUE_index(ByVal queue As Long) As Long
    NETQUEUE_index = nettypes_c.NETQUEUE_index(queue)
End Function

Public Function NETQUEUE_queueType(ByVal queue As Long) As Long
    NETQUEUE_queueType = nettypes_c.NETQUEUE_queueType(queue)
End Function

Public Function NETQUEUE_exclude(ByVal queue As Long) As Long
    NETQUEUE_exclude = nettypes_c.NETQUEUE_exclude(queue)
End Function

Public Function NETQUEUE_isPair(ByVal queue As Long) As Boolean
    NETQUEUE_isPair = nettypes_c.NETQUEUE_isPair(queue)
End Function

Public Function NETnetTmpQueue(ByVal tmpPlayer As Long) As Long
    NETnetTmpQueue = nettypes_c.NETnetTmpQueue(tmpPlayer)
End Function

Public Function NETnetQueue(ByVal player As Long, Optional ByVal excludePlayer As Long = 255) As Long
    NETnetQueue = nettypes_c.NETnetQueue(player, excludePlayer)
End Function

Public Function NETgameQueue(ByVal player As Long) As Long
    NETgameQueue = nettypes_c.NETgameQueue(player)
End Function

Public Function NETgameQueueForced(ByVal player As Long) As Long
    NETgameQueueForced = nettypes_c.NETgameQueueForced(player)
End Function

Public Function NETbroadcastQueue(Optional ByVal excludePlayer As Long = 255) As Long
    NETbroadcastQueue = nettypes_c.NETbroadcastQueue(excludePlayer)
End Function

Public Sub NETinitQueue(ByVal queue As Long)
    nettypes_c.NETinitQueue queue
End Sub

Public Sub NETsetNoSendOverNetwork(ByVal queue As Long)
    nettypes_c.NETsetNoSendOverNetwork queue
End Sub

Public Sub NETmoveQueue(ByVal src As Long, ByVal dst As Long)
    nettypes_c.NETmoveQueue src, dst
End Sub

Public Sub NETdeleteQueue()
    nettypes_c.NETdeleteQueue
End Sub

' Dump raw data from sockets and raw data sent via host here: data(0 .. dataLen - 1).
Public Sub NETinsertRawData(ByVal queue As Long, data() As Byte, ByVal dataLen As Long)
    nettypes_c.NETinsertRawData queue, data, dataLen
End Sub

Public Sub NETinsertMessageFromNet(ByVal queue As Long, ByVal newMessage As NetMessage)
    nettypes_c.NETinsertMessageFromNet queue, newMessage
End Sub

Public Function NETisMessageReady(ByVal queue As Long) As Boolean
    NETisMessageReady = nettypes_c.NETisMessageReady(queue)
End Function

Public Function NETincompleteMessageDataBuffered(ByVal queue As Long) As Long
    NETincompleteMessageDataBuffered = nettypes_c.NETincompleteMessageDataBuffered(queue)
End Function

Public Function NETgetMessage(ByVal queue As Long) As NetMessage
    Set NETgetMessage = nettypes_c.NETgetMessage(queue)
End Function

Public Sub NETpop(ByVal queue As Long)
    nettypes_c.NETpop queue
End Sub

Public Sub NETflushGameQueues()
    nettypes_c.NETflushGameQueues
End Sub

Public Function NETgameIsBehindPlayersByAtLeast(Optional ByVal numGameTimeUpdates As Long = 2) As Boolean
    NETgameIsBehindPlayersByAtLeast = nettypes_c.NETgameIsBehindPlayersByAtLeast(numGameTimeUpdates)
End Function

' (VB) the joining screen's tmpJoiningQueuePair (QUEUE_TRANSIENT_JOIN's pair, which NETinitQueue makes):
' &tmpJoiningQueuePair->send's numMessagesForNet / getMessageForNet / popMessageForNet; delete tmpJoiningQueuePair
Public Function NETtransientJoinSendQueue_numMessagesForNet() As Long
    NETtransientJoinSendQueue_numMessagesForNet = nettypes_c.NETtransientJoinSendQueue_numMessagesForNet()
End Function

Public Function NETtransientJoinSendQueue_getMessageForNet() As NetMessage
    Set NETtransientJoinSendQueue_getMessageForNet = nettypes_c.NETtransientJoinSendQueue_getMessageForNet()
End Function

Public Sub NETtransientJoinSendQueue_popMessageForNet()
    nettypes_c.NETtransientJoinSendQueue_popMessageForNet
End Sub

Public Sub NETtransientJoinQueuePair_delete()
    nettypes_c.NETtransientJoinQueuePair_delete
End Sub

' tmpJoiningQueuePair != nullptr
Public Function NETtransientJoinQueuePair_exists() As Boolean
    NETtransientJoinQueuePair_exists = nettypes_c.NETtransientJoinQueuePair_exists()
End Function

' MessageReader::msgType (a Friend property of the class: the exe reads it here)
Public Function MessageReader_msgType(ByVal r As MessageReader) As Long
    MessageReader_msgType = r.msgType
End Function

' MessageReader(NetMessage const &msg)
Public Function MessageReader_new(ByVal message As NetMessage) As MessageReader
    Set MessageReader_new = netqueue_c.MessageReader_new(message)
End Function

' void NETsetSessionKeys(uint8_t player, SessionKeys&& keys): the keys as their bytes (whole arrays)
Public Sub NETsetSessionKeys(ByVal player As Long, receiveKey() As Byte, sendKey() As Byte)
    nettypes_c.NETsetSessionKeys player, receiveKey, sendKey
End Sub

Public Sub NETclearSessionKeys()
    nettypes_c.NETclearSessionKeys
End Sub

' void NETclearSessionKeys(uint8_t player)
Public Sub NETclearSessionKeys_player(ByVal player As Long)
    nettypes_c.NETclearSessionKeys_player player
End Sub

Public Function NETisExpectedSecuredMessageType(ByVal typ As Long) As Boolean
    NETisExpectedSecuredMessageType = nettypes_c.NETisExpectedSecuredMessageType(typ)
End Function

' std::vector<uint8_t> output: output(0 .. outputSize - 1), appended to
Public Sub NETbytesOutputToVector(data() As Byte, ByVal dataSize As Long, output() As Byte, outputSize As Long)
    nettypes_c.NETbytesOutputToVector data, dataSize, output, outputSize
End Sub

Public Function NETisReplay() As Boolean
    NETisReplay = nettypes_c.NETisReplay()
End Function

Public Function NETloadReplay(ByVal filename As String, ByVal optionsHandler As ReplayOptionsHandler) As Boolean
    NETloadReplay = nettypes_c.NETloadReplay(filename, optionsHandler)
End Function

Public Sub NETshutdownReplay()
    nettypes_c.NETshutdownReplay
End Sub

'--------------------------------------------------------------------------
' netreplay.h

Public Function NETreplaySaveStart(ByVal subdir As String, ByVal optionsHandler As ReplayOptionsHandler, ByVal maxReplaysSaved As Long, ByVal appendPlayerToFilename As Boolean) As String
    NETreplaySaveStart = netreplay_c.NETreplaySaveStart(subdir, optionsHandler, maxReplaysSaved, appendPlayerToFilename)
End Function

Public Function NETreplaySaveStop(ByVal optionsHandler As ReplayOptionsHandler) As Boolean
    NETreplaySaveStop = netreplay_c.NETreplaySaveStop(optionsHandler)
End Function

Public Sub NETreplaySaveNetMessage(ByVal message As NetMessage, ByVal player As Long)
    netreplay_c.NETreplaySaveNetMessage message, player
End Sub

Public Function NETreplayLoadStart(ByVal filename As String, ByVal optionsHandler As ReplayOptionsHandler, output_replayFormatVer As Long) As Boolean
    NETreplayLoadStart = netreplay_c.NETreplayLoadStart(filename, optionsHandler, output_replayFormatVer)
End Function

Public Function NETreplayLoadNetMessage(message As NetMessage, player As Long) As Boolean
    NETreplayLoadNetMessage = netreplay_c.NETreplayLoadNetMessage(message, player)
End Function

Public Function NETreplayLoadStop() As Boolean
    NETreplayLoadStop = netreplay_c.NETreplayLoadStop()
End Function

'--------------------------------------------------------------------------
' nettypes.h: serialising. Each NET* takes the MessageReader or the
' MessageWriter As Object (one procedure for C++'s overload pair) and the
' value ByRef: a reader sets it, a writer reads it.

Public Function NETbeginEncode(ByVal queue As Long, ByVal typ As Long) As MessageWriter
    Set NETbeginEncode = nettypes_c.NETbeginEncode(queue, typ)
End Function

Public Function NETbeginEncodeSecured(ByVal queue As Long, ByVal typ As Long) As MessageWriter
    Set NETbeginEncodeSecured = nettypes_c.NETbeginEncodeSecured(queue, typ)
End Function

Public Function NETbeginDecodeSecured(ByVal queue As Long, ByVal typ As Long) As MessageReader
    Set NETbeginDecodeSecured = nettypes_c.NETbeginDecodeSecured(queue, typ)
End Function

Public Function NETbeginDecode(ByVal queue As Long, ByVal typ As Long) As MessageReader
    Set NETbeginDecode = nettypes_c.NETbeginDecode(queue, typ)
End Function

Public Function NETend(ByVal c As Object) As Boolean
    NETend = nettypes_c.NETend(c)
End Function

Public Sub NETuint8_t(ByVal c As Object, v As Long)
    nettypes_c.NETuint8_t c, v
End Sub

Public Sub NETint8_t(ByVal c As Object, v As Long)
    nettypes_c.NETint8_t c, v
End Sub

Public Sub NETuint16_t(ByVal c As Object, v As Long)
    nettypes_c.NETuint16_t c, v
End Sub

Public Sub NETint16_t(ByVal c As Object, v As Long)
    nettypes_c.NETint16_t c, v
End Sub

Public Sub NETuint32_t(ByVal c As Object, v As Long)
    nettypes_c.NETuint32_t c, v
End Sub

Public Sub NETint32_t(ByVal c As Object, v As Long)
    nettypes_c.NETint32_t c, v
End Sub

' 64-bit values are Doubles (exact below 2^53)
Public Sub NETuint64_t(ByVal c As Object, v As Double)
    nettypes_c.NETuint64_t c, v
End Sub

Public Sub NETint64_t(ByVal c As Object, v As Double)
    nettypes_c.NETint64_t c, v
End Sub

Public Sub NETbool(ByVal c As Object, v As Boolean)
    nettypes_c.NETbool c, v
End Sub

' template NETenum: a uint32_t
Public Sub NETenum(ByVal c As Object, v As Long)
    nettypes_c.NETenum c, v
End Sub

Public Sub NETwzstring(ByVal c As Object, s As String)
    nettypes_c.NETwzstring c, s
End Sub

' NETstring(r, char *str, uint16_t maxlen) / NETstring(w, const char *str, uint16_t maxlen)
Public Sub NETstring(ByVal c As Object, s As String, ByVal maxlen As Long)
    nettypes_c.NETstring c, s, maxlen
End Sub

' NETstring(r, std::string &s, uint32_t maxLen) / NETstring(w, const std::string &s, uint32_t maxLen)
Public Sub NETstring_std(ByVal c As Object, s As String, Optional ByVal maxLen As Long = 65536)
    nettypes_c.NETstring_std c, s, maxLen
End Sub

Public Sub NETbin(ByVal c As Object, b() As Byte, ByVal len_ As Long)
    nettypes_c.NETbin c, b, len_
End Sub

Public Sub NETbytes(ByVal c As Object, b() As Byte, Optional ByVal maxLen As Long = 10000)
    nettypes_c.NETbytes c, b, maxLen
End Sub

Public Sub NETnetMessage(ByVal c As Object, msg As NetMessage)
    nettypes_c.NETnetMessage c, msg
End Sub

' (cluster J part 2) the exe's lobby hosting handler (netlobby.cpp) calling back into netplay: the eventCallbacks of
' NETcreateLobbyHostingHandler (netlobby_c) and checkJoinRequest's result handler (hasResult: the CheckJoinResult
' has a value, pass its CheckJoinRequestResult::pass; otherwise errCode and errMessage are its LobbyError's)
Public Sub NETlobbyGameCreatedCallback(ByVal gameId As String)
    netlobby_c.lobbyGameCreatedCallback gameId
End Sub

Public Sub NETlobbyGamePublishedCallback(ByVal gameId As String, ByVal motd As String)
    netlobby_c.lobbyGamePublishedCallback gameId, motd
End Sub

Public Sub NETlobbyGameGoneCallback(ByVal hasError As Boolean, ByVal errCode As String, ByVal errMessage As String)
    netlobby_c.lobbyGameGoneCallback hasError, errCode, errMessage
End Sub

Public Sub NETtempSocketHandleLobbyJoinCheckResult(ByVal i As Long, ByVal uniqueJoinID As String, ByVal hasResult As Boolean, _
        ByVal pass As Boolean, ByVal errCode As String, ByVal errMessage As String)
    netplay_conn_c.NETtempSocketHandleLobbyJoinCheckResult i, uniqueJoinID, hasResult, pass, errCode, errMessage
End Sub

' (cluster J part 2) curl's part of the exe's urlrequest (urltransfer_c: HTTP/1.1 over Winsock and Schannel). method,
' url, headers ("Name: value" lines ending in vbCrLf), body(0 .. bodyLen - 1) (bodyLen -1: none), protocol
' (InternetProtocol: 0 IP_ANY, 1 IPv4, 2 IPv6), connectTimeoutMs (0: curl's default), followRedirects, userAgent,
' maxFileSize (0: none): the transfer's handle
Public Function NETurlTransferStart(ByVal method As String, ByVal url As String, ByVal headers As String, body() As Byte, _
        ByVal bodyLen As Long, ByVal protocol As Long, ByVal connectTimeoutMs As Long, ByVal followRedirects As Boolean, _
        ByVal userAgent As String, ByVal maxFileSize As Double) As Long
    NETurlTransferStart = urltransfer_c.urlTransferStart(method, url, headers, body, bodyLen, protocol, connectTimeoutMs, _
        followRedirects, userAgent, maxFileSize)
End Function

' True when the transfer is done (waitMs: how long it may wait for its socket when nothing can be done)
Public Function NETurlTransferPoll(ByVal handle As Long, ByVal waitMs As Long) As Boolean
    NETurlTransferPoll = urltransfer_c.urlTransferPoll(handle, waitMs)
End Function

' the CURLcode, its text, CURLINFO_RESPONSE_CODE, CURLINFO_PRIMARY_IP, CURLINFO_RETRY_AFTER
Public Function NETurlTransferResult(ByVal handle As Long, resultText As String, httpStatus As Long, primaryIp As String, _
        retryAfter As Double) As Long
    NETurlTransferResult = urltransfer_c.urlTransferResult(handle, resultText, httpStatus, primaryIp, retryAfter)
End Function

' the header callback's lines (each ending in vbLf)
Public Function NETurlTransferHeaders(ByVal handle As Long) As String
    NETurlTransferHeaders = urltransfer_c.urlTransferHeaders(handle)
End Function

' the body: body(0 .. n - 1), n returned
Public Function NETurlTransferBody(ByVal handle As Long, body() As Byte) As Long
    NETurlTransferBody = urltransfer_c.urlTransferBody(handle, body)
End Function

' the transfer closed (aborted when running) and its handle freed
Public Sub NETurlTransferClose(ByVal handle As Long)
    urltransfer_c.urlTransferClose handle
End Sub

' curl_version()'s place
Public Function NETurlTransferVersion() As String
    NETurlTransferVersion = urltransfer_c.urlTransferVersion()
End Function
