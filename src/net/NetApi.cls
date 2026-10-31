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
' modules; this class only forwards to them. Calls the other way go through
' INetGame.
Option Explicit

' connection_provider_registry.h: enum class ConnectionProviderType : uint8_t (no GNS backend)
Public Enum ConnectionProviderType
    ConnectionProviderType_TCP_DIRECT = 0
End Enum

' netlobby.h: enum class HostJoinOptionValue
Public Enum HostJoinOptionValue
    HostJoinOptionValue_BlockAll = 0
    HostJoinOptionValue_BlockSuspicious = 1
    HostJoinOptionValue_Allow = 2
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

' netplay.h: enum CONNECTION_STATUS
Public Enum CONNECTION_STATUS
    CONNECTIONSTATUS_PLAYER_DROPPED = 0
    CONNECTIONSTATUS_PLAYER_LEAVING = 1
    CONNECTIONSTATUS_DESYNC = 2
    CONNECTIONSTATUS_WAITING_FOR_PLAYER = 3

    CONNECTIONSTATUS_NORMAL = 4
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
' netplay.h's NETPLAY NetPlay: its fields as NetPlay_<field>

Public Property Get NetPlay_bComms() As Boolean
    NetPlay_bComms = stubs_net.NetPlay.bComms
End Property

Public Property Let NetPlay_bComms(ByVal value As Boolean)
    stubs_net.NetPlay.bComms = value
End Property

' NetPlay.playercount: uint32_t, the number of players in the game (read only here)
Public Property Get NetPlay_playercount() As Long
    NetPlay_playercount = netplay_c.NetPlay_get_playercount()
End Property

Public Property Get NetPlay_isHost() As Boolean
    NetPlay_isHost = stubs_net.NetPlay.isHost
End Property

Public Property Let NetPlay_isHost(ByVal value As Boolean)
    stubs_net.NetPlay.isHost = value
End Property

Public Property Get NetPlay_hostPlayer() As Long
    NetPlay_hostPlayer = stubs_net.NetPlay.hostPlayer
End Property

Public Property Let NetPlay_hostPlayer(ByVal value As Long)
    stubs_net.NetPlay.hostPlayer = value
End Property

Public Property Get NetPlay_players_allocated(ByVal i As Long) As Boolean
    NetPlay_players_allocated = stubs_net.NetPlay_players_allocated(i)
End Property

Public Property Let NetPlay_players_allocated(ByVal i As Long, ByVal value As Boolean)
    stubs_net.set_NetPlay_players_allocated i, value
End Property

Public Property Get NetPlay_players_isSpectator(ByVal i As Long) As Boolean
    NetPlay_players_isSpectator = stubs_net.NetPlay_players_isSpectator(i)
End Property

Public Property Let NetPlay_players_isSpectator(ByVal i As Long, ByVal value As Boolean)
    stubs_net.set_NetPlay_players_isSpectator i, value
End Property

' NetPlay.players[i].faction (a FactionID)
Public Property Get NetPlay_players_faction(ByVal i As Long) As Long
    NetPlay_players_faction = stubs_net.NetPlay_players_faction(i)
End Property

Public Property Let NetPlay_players_faction(ByVal i As Long, ByVal value As Long)
    stubs_net.set_NetPlay_players_faction i, value
End Property

' NetPlay.players[i].position (int32_t: the map starting position)
Public Property Get NetPlay_players_position(ByVal i As Long) As Long
    NetPlay_players_position = stubs_net.NetPlay_players_position(i)
End Property

Public Property Let NetPlay_players_position(ByVal i As Long, ByVal value As Long)
    stubs_net.set_NetPlay_players_position i, value
End Property

' NetPlay.players[i].difficulty (an AIDifficulty: -2 HUMAN, -1 DISABLED, 0 SUPEREASY .. 4 INSANE)
Public Property Get NetPlay_players_difficulty(ByVal i As Long) As Long
    NetPlay_players_difficulty = stubs_net.NetPlay_players_difficulty(i)
End Property

Public Property Let NetPlay_players_difficulty(ByVal i As Long, ByVal value As Long)
    stubs_net.set_NetPlay_players_difficulty i, value
End Property

' NetPlay.players[i].name (char[StringSize])
Public Property Get NetPlay_players_name(ByVal i As Long) As String
    NetPlay_players_name = stubs_net.NetPlay_players_name(i)
End Property

Public Property Let NetPlay_players_name(ByVal i As Long, ByVal value As String)
    stubs_net.set_NetPlay_players_name i, value
End Property

' NetPlay.players[i].isAdmin
Public Property Get NetPlay_players_isAdmin(ByVal i As Long) As Boolean
    NetPlay_players_isAdmin = stubs_net.NetPlay_players_isAdmin(i)
End Property

Public Property Let NetPlay_players_isAdmin(ByVal i As Long, ByVal value As Boolean)
    stubs_net.set_NetPlay_players_isAdmin i, value
End Property

Public Property Get NetPlay_isHostAlive() As Boolean
    NetPlay_isHostAlive = stubs_net.NetPlay.isHostAlive
End Property

Public Property Let NetPlay_isHostAlive(ByVal value As Boolean)
    stubs_net.NetPlay.isHostAlive = value
End Property

' NetPlay.players.size(): MAX_CONNECTED_PLAYERS, + 1 with a replay's spectator
Public Function NetPlay_players_size() As Long
    NetPlay_players_size = stubs_net.NetPlay_players_size()
End Function

' NetPlay.players.push_back(PLAYER())
Public Sub NetPlay_players_push_back()
    stubs_net.NetPlay_players_push_back
End Sub

' NetPlay.players.resize(n)
Public Sub NetPlay_players_resize(ByVal n As Long)
    stubs_net.NetPlay_players_resize n
End Sub

' int32_t: Which colour slot this player is using
Public Property Get NetPlay_players_colour(ByVal i As Long) As Long
    NetPlay_players_colour = stubs_net.NetPlay_players_colour(i)
End Property

Public Property Let NetPlay_players_colour(ByVal i As Long, ByVal value As Long)
    stubs_net.set_NetPlay_players_colour i, value
End Property

' uint32_t: Time cardiac arrest started
Public Property Get NetPlay_players_heartattacktime(ByVal i As Long) As Long
    NetPlay_players_heartattacktime = stubs_net.NetPlay_players_heartattacktime(i)
End Property

Public Property Let NetPlay_players_heartattacktime(ByVal i As Long, ByVal value As Long)
    stubs_net.set_NetPlay_players_heartattacktime i, value
End Property

' If connection is still alive or not
Public Property Get NetPlay_players_heartbeat(ByVal i As Long) As Boolean
    NetPlay_players_heartbeat = stubs_net.NetPlay_players_heartbeat(i)
End Property

Public Property Let NetPlay_players_heartbeat(ByVal i As Long, ByVal value As Boolean)
    stubs_net.set_NetPlay_players_heartbeat i, value
End Property

' int32_t: Which team we are on (int32_t::max for spectator team)
Public Property Get NetPlay_players_team(ByVal i As Long) As Long
    NetPlay_players_team = stubs_net.NetPlay_players_team(i)
End Property

Public Property Let NetPlay_players_team(ByVal i As Long, ByVal value As Long)
    stubs_net.set_NetPlay_players_team i, value
End Property

' player ready to start?
Public Property Get NetPlay_players_ready(ByVal i As Long) As Boolean
    NetPlay_players_ready = stubs_net.NetPlay_players_ready(i)
End Property

Public Property Let NetPlay_players_ready(ByVal i As Long, ByVal value As Boolean)
    stubs_net.set_NetPlay_players_ready i, value
End Property

' int8_t: index into sorted list of AIs, zero is always default AI
Public Property Get NetPlay_players_ai(ByVal i As Long) As Long
    NetPlay_players_ai = stubs_net.NetPlay_players_ai(i)
End Property

Public Property Let NetPlay_players_ai(ByVal i As Long, ByVal value As Long)
    stubs_net.set_NetPlay_players_ai i, value
End Property

' if we are running a autogame (AI controls us)
Public Property Get NetPlay_players_autoGame(ByVal i As Long) As Boolean
    NetPlay_players_autoGame = stubs_net.NetPlay_players_autoGame(i)
End Property

Public Property Let NetPlay_players_autoGame(ByVal i As Long, ByVal value As Boolean)
    stubs_net.set_NetPlay_players_autoGame i, value
End Property

Public Property Get NetPlay_isPortMappingEnabled() As Boolean
    NetPlay_isPortMappingEnabled = stubs_net.NetPlay.isPortMappingEnabled
End Property

Public Property Let NetPlay_isPortMappingEnabled(ByVal value As Boolean)
    stubs_net.NetPlay.isPortMappingEnabled = value
End Property

'--------------------------------------------------------------------------
' netplay.cpp's settings, netlobby.h

Public Property Get netGameserverPortOverride() As Boolean
    netGameserverPortOverride = stubs_net.netGameserverPortOverride
End Property

Public Property Let netGameserverPortOverride(ByVal value As Boolean)
    stubs_net.netGameserverPortOverride = value
End Property

Public Property Get iptoconnect() As String
    iptoconnect = stubs_net.iptoconnect
End Property

Public Property Let iptoconnect(ByVal value As String)
    stubs_net.iptoconnect = value
End Property

Public Property Get cliConnectAsSpectator() As Boolean
    cliConnectAsSpectator = stubs_net.cliConnectAsSpectator
End Property

Public Property Let cliConnectAsSpectator(ByVal value As Boolean)
    stubs_net.cliConnectAsSpectator = value
End Property

Public Sub NET_InitPlayer(ByVal i As Long, ByVal initPosition As Boolean, Optional ByVal initTeams As Boolean = False, Optional ByVal initSpectator As Boolean = False)
    netplay_c.NET_InitPlayer i, initPosition, initTeams, initSpectator
End Sub

Public Sub NET_InitPlayers(Optional ByVal initTeams As Boolean = False, Optional ByVal initSpectator As Boolean = False)
    netplay_c.NET_InitPlayers initTeams, initSpectator
End Sub

Public Sub NETsetAsyncJoinApprovalRequired(ByVal enabled As Boolean)
    stubs_net.NETsetAsyncJoinApprovalRequired enabled
End Sub

Public Function netlobby_GetDefaultLobbyAddress() As String
    netlobby_GetDefaultLobbyAddress = stubs_net.netlobby_GetDefaultLobbyAddress()
End Function

Public Sub NETsetLobbyserverAddress(ByVal lobbyAddress As String)
    stubs_net.NETsetLobbyserverAddress lobbyAddress
End Sub

Public Function NETgetLobbyserverAddress() As String
    NETgetLobbyserverAddress = stubs_net.NETgetLobbyserverAddress()
End Function

Public Sub NETsetGameserverPort(ByVal port As Long)
    stubs_net.NETsetGameserverPort port
End Sub

Public Function NETgetGameserverPort() As Long
    NETgetGameserverPort = stubs_net.NETgetGameserverPort()
End Function

Public Sub NETsetJoinPreferenceIPv6(ByVal bTryIPv6First As Boolean)
    stubs_net.NETsetJoinPreferenceIPv6 bTryIPv6First
End Sub

Public Function NETgetJoinPreferenceIPv6() As Boolean
    NETgetJoinPreferenceIPv6 = stubs_net.NETgetJoinPreferenceIPv6()
End Function

Public Sub NETsetDefaultMPHostFreeChatPreference(ByVal enabled As Boolean)
    stubs_net.NETsetDefaultMPHostFreeChatPreference enabled
End Sub

Public Function NETgetDefaultMPHostFreeChatPreference() As Boolean
    NETgetDefaultMPHostFreeChatPreference = stubs_net.NETgetDefaultMPHostFreeChatPreference()
End Function

Public Sub NETsetEnableTCPNoDelay(ByVal enabled As Boolean)
    stubs_net.NETsetEnableTCPNoDelay enabled
End Sub

Public Function NETgetEnableTCPNoDelay() As Boolean
    NETgetEnableTCPNoDelay = stubs_net.NETgetEnableTCPNoDelay()
End Function

' C: netlobby::HostJoinOptions NETgetLobbyHostJoinOptions(): its three fields
Public Sub NETgetLobbyHostJoinOptions(botProt As Long, proxyIPs As Long, hostingIPs As Long)
    stubs_net.NETgetLobbyHostJoinOptions botProt, proxyIPs, hostingIPs
End Sub

Public Sub NETsetLobbyHostJoinOptions(ByVal botProt As Long, ByVal proxyIPs As Long, ByVal hostingIPs As Long)
    stubs_net.NETsetLobbyHostJoinOptions botProt, proxyIPs, hostingIPs
End Sub

' C: SpectatorInfo SpectatorInfo::currentNetPlayState(), packed as its
' toUint32(): spectatorsJoined in the high 16 bits, totalSpectatorSlots in
' the low 16
Public Function SpectatorInfo_currentNetPlayState() As Long
    SpectatorInfo_currentNetPlayState = stubs_net.SpectatorInfo_currentNetPlayState()
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
' stand-ins (stubs_net): netplay.h

Public Function NETclose() As Long
    NETclose = netplay_c.NETclose()
End Function

Public Function netplayShutDown() As Boolean
    netplayShutDown = stubs_net.netplayShutDown()
End Function

'--------------------------------------------------------------------------
' nettypes.h, netqueue.h: the queues. A NETQUEUE is a Long:
' queueType * &H10000 + index * &H100 + exclude.

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

Public Sub NETdeleteQueue()
    nettypes_c.NETdeleteQueue
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

'--------------------------------------------------------------------------
' netplay.h: the local half of the message traffic

Public Function NETrecvNet(queue As Long, typ As Long) As Boolean
    NETrecvNet = netplay_c.NETrecvNet(queue, typ)
End Function

Public Function NETrecvGame(queue As Long, typ As Long) As Boolean
    NETrecvGame = netplay_c.NETrecvGame(queue, typ)
End Function

Public Sub NETflush()
    netplay_c.NETflush
End Sub

Public Function messageTypeToString(ByVal messageType_ As Long) As String
    messageTypeToString = netplay_c.messageTypeToString(messageType_)
End Function

Public Sub NETsetPlayerConnectionStatus(ByVal status As Long, ByVal player As Long)
    netplay_c.NETsetPlayerConnectionStatus status, player
End Sub

Public Function NETcheckPlayerConnectionStatus(ByVal status As Long, ByVal player As Long) As Boolean
    NETcheckPlayerConnectionStatus = netplay_c.NETcheckPlayerConnectionStatus(status, player)
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
' netplay.h: the host's start, the players' info messages, the game's lock

' bool netPlayersUpdated (update flags)
Public Property Get netPlayersUpdated() As Boolean
    netPlayersUpdated = netplay_c.netPlayersUpdated
End Property

Public Property Let netPlayersUpdated(ByVal value As Boolean)
    netplay_c.netPlayersUpdated = value
End Property

' C: bool NEThostGame(const char *SessionName, const char *PlayerName, const EcKey& playerIdentity, bool spectatorHost,
' uint32_t gameType, uint32_t blindMode, UDWORD plyrs, uint16_t defaultOpenSpectatorSlots, uint8_t alliancesType,
' uint8_t techLevel, uint8_t powerLevel, uint8_t basesLevel), without the playerIdentity (the connections', cluster J)
Public Function NEThostGame(ByVal SessionName As String, ByVal PlayerName As String, ByVal spectatorHost As Boolean, _
                            ByVal gameType As Long, ByVal blindMode As Long, ByVal plyrs As Long, ByVal defaultOpenSpectatorSlots As Long, _
                            ByVal alliancesType As Long, ByVal techLevel As Long, ByVal powerLevel As Long, ByVal basesLevel As Long) As Boolean
    NEThostGame = netplay_c.NEThostGame(SessionName, PlayerName, spectatorHost, gameType, blindMode, plyrs, defaultOpenSpectatorSlots, _
                                        alliancesType, techLevel, powerLevel, basesLevel)
End Function

' Stop the dplay interface from accepting more players.
Public Function NEThaltJoining(ByVal matchStarted As Boolean) As Boolean
    NEThaltJoining = netplay_c.NEThaltJoining(matchStarted)
End Function

Public Sub NETBroadcastPlayerInfo(ByVal index As Long)
    netplay_c.NETBroadcastPlayerInfo index
End Sub

Public Sub NETBroadcastTwoPlayerInfo(ByVal index1 As Long, ByVal index2 As Long)
    netplay_c.NETBroadcastTwoPlayerInfo index1, index2
End Sub

Public Sub NETSendAllPlayerInfoTo(ByVal to_ As Long)
    netplay_c.NETSendAllPlayerInfoTo to_
End Sub

Public Function NETGameIsLocked() As Boolean
    NETGameIsLocked = netplay_c.NETGameIsLocked()
End Function

Public Function NET_getLobbyDisabled() As Boolean
    NET_getLobbyDisabled = netplay_c.NET_getLobbyDisabled()
End Function

' Sets if the game is password protected or not
Public Sub NETGameLocked(ByVal flag As Boolean)
    netplay_c.NETGameLocked flag
End Sub

' Sets the game password
Public Sub NETsetGamePassword(ByVal password As String)
    netplay_c.NETsetGamePassword password
End Sub

' ingameLimitFlags: ingame.flags (uint8_t)
Public Sub NETsetLobbyLimitFlags(ByVal ingameLimitFlags As Long)
    netplay_c.NETsetLobbyLimitFlags ingameLimitFlags
End Sub
