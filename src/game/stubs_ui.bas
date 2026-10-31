Attribute VB_Name = "stubs_ui"
' Clusters F and G's merge: what the widgets, the front end and the in-game interface call that
' no module provides yet (load and save, the campaign, the in-game multiplayer menu, videos, the
' network lobby, ...). Stand-ins keep the C++ name and signature; a few are real ports of a function
' of a module not ported yet, and say where they belong.
Option Explicit

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS


' src/transporter.h: #define IDTRANS_LAUNCH 9010 //The Transporter Launch button

' lib/netplay/netplay.h: enum NetStatisticType
Public Const NetStatisticRawBytes As Long = 0

Public Const NetStatisticUncompressedBytes As Long = 1

Public Const NetStatisticPackets As Long = 2






' #define MULTIMENU_FORM MULTIMENU
Public Const MULTIMENU_FORM As Long = MULTIMENU


' src/loadsave.cpp: extern bool lastSaveMP; (the last save was a skirmish/multiplayer one)

' src/notifications.h (cluster G): enum class WZ_Notification_Dismissal_Reason
Public Enum WZ_Notification_Dismissal_Reason
    WZ_Notification_Dismissal_Reason_USER_DISMISSED
    WZ_Notification_Dismissal_Reason_ACTION_BUTTON_CLICK
    WZ_Notification_Dismissal_Reason_PROGRAMMATIC
End Enum

' src/notifications.h (cluster G): enum class NotificationScope
Public Enum NotificationScope
    NotificationScope_DISPLAYED_ONLY
    NotificationScope_QUEUED_ONLY
    NotificationScope_DISPLAYED_AND_QUEUED
End Enum

' src/notifications.h (cluster G): struct WZ_Notification, flattened: what the front end sets.
' largeIcon: WZ_Notification_Image's file path ("" = none); action_title / action_onAction:
' WZ_Notification_Action (onAction: void (const WZ_Notification&), Nothing = none); onDismissed:
' void (const WZ_Notification&, WZ_Notification_Dismissal_Reason) (the WzFunc gets the reason as its
' second argument); displayOptions_*: WZ_Notification_Display_Options (identifier "" = default;
' makeIgnorable: numTimesSeenBeforeDoNotShowAgainOption; makeOneTime: oneTime).
Public Type WZ_Notification
    duration As Long                    ' uint32_t: set to 0 for "until dismissed by user"
    contentTitle As String
    contentText As String
    largeIcon As String
    action_title As String
    action_onAction As WzFunc
    onDismissed As WzFunc
    tag As String
    displayOptions_identifier As String
    displayOptions_numTimesSeenBeforeDoNotShowAgainOption As Long
    displayOptions_oneTime As Boolean
End Type

' lib/netplay/netplay.h: enum class SpectatorToPlayerMoveResult
Public Enum SpectatorToPlayerMoveResult
    SpectatorToPlayerMoveResult_SUCCESS
    SpectatorToPlayerMoveResult_NEEDS_SLOT_SELECTION
    SpectatorToPlayerMoveResult_FAILED
End Enum

' struct screeninfo
Public Type screeninfo
    width As Long
    height As Long
    pixel_density As Single
    refresh_rate As Single
    screen As Long                          ' uint32_t
End Type

' enum class MinimizeOnFocusLossBehavior
Public Const MinimizeOnFocusLossBehavior_Auto As Long = -1

Public Const MinimizeOnFocusLossBehavior_Off As Long = 0

Public Const MinimizeOnFocusLossBehavior_On_Fullscreen As Long = 1

' enum class KeyMappingSlot



' enum class KeyMappingType



' enum class KeyMappingInputSource


' enum class KeyAction



' namespace netlobby: enum class HostJoinOptionType
Public Enum HostJoinOptionType
    HostJoinOptionType_BotProt
    HostJoinOptionType_ProxyIPs
    HostJoinOptionType_HostingIPs
End Enum

' netplay.h: enum class LobbyHostDirectJoinOption
Public Enum LobbyHostDirectJoinOption
    LobbyHostDirectJoinOption_None
    LobbyHostDirectJoinOption_Local
    LobbyHostDirectJoinOption_All
End Enum

' netplay.h: enum class NET_LOBBY_OPT_FIELD (WzMultiplayerOptionsTitleUI.cls uses NET_LOBBY_OPT_FIELD_MAPNAME too)
Public Enum NET_LOBBY_OPT_FIELD
    NET_LOBBY_OPT_FIELD_INVALID
    NET_LOBBY_OPT_FIELD_GNAME
    NET_LOBBY_OPT_FIELD_MAPNAME
    NET_LOBBY_OPT_FIELD_HOSTNAME
    NET_LOBBY_OPT_FIELD_MAX
End Enum

' netplay.h: #define password_string_size 64 (netplay_c's copy is Private)
Public Const password_string_size As Long = 64

' multiplay.h: enum TechLevel
Public Const TECH_1 As Long = 1

Public Const TECH_2 As Long = 2

Public Const TECH_3 As Long = 3

Public Const TECH_4 As Long = 4

' (stand-in state) NETgetLobbyHostJoinOptions / NETgetLobbyHostDirectJoinOption
Private hostJoinOpts_set As Boolean

Private hostJoinOpts_botProt As Long

Private hostJoinOpts_proxyIPs As Long

Private hostJoinOpts_hostingIPs As Long

Private lobbyDirectJoinOption_set As Boolean

Private lobbyDirectJoinOption As Long

' src/hci.cpp: bool createReplayControllerOverlay() (S: ReplayControllerWidget is not ported; replays only)
Public Function createReplayControllerOverlay() As Boolean
    createReplayControllerOverlay = False
End Function




' src/screens/helpscreen.cpp: std::shared_ptr<W_HELP_OVERLAY_SCREEN> W_HELP_OVERLAY_SCREEN::make(const OnCloseFunc& onCloseFunc)
' (S: Nothing; intShowWidgetHelp then shows nothing)
Public Function W_HELP_OVERLAY_SCREEN_make(ByVal onCloseFunc As WzFunc) As W_SCREEN
    Set W_HELP_OVERLAY_SCREEN_make = Nothing
End Function

' src/screens/helpscreen.cpp: void W_HELP_OVERLAY_SCREEN::setHelpFromWidgets(const std::shared_ptr<WIDGET>& root) (S)
Public Sub W_HELP_OVERLAY_SCREEN_setHelpFromWidgets(ByVal screen As W_SCREEN, ByVal root As WIDGET)
End Sub

' src/multimenu.cpp: bool intCloseMultiMenu() (S: the in-game multiplayer menu is not ported)
Public Function intCloseMultiMenu() As Boolean
    intCloseMultiMenu = True
End Function

' src/multimenu.cpp: void intCloseMultiMenuNoAnim() (S)
Public Sub intCloseMultiMenuNoAnim()
End Sub








' src/display.cpp: startDeliveryPosition(FLAG_POSITION *psFlag) for a flag that is not in aFlagPos (the debug menu's
' local FLAG_POSITION) (S: nothing; display_c's takes an aFlagPos handle)
Public Sub startDeliveryPosition_flag(psFlag As FLAG_POSITION)
End Sub

' src/multistruct.cpp: bool SendBuildFinished(STRUCTURE *psStruct) for a STRUCTURE that is not in the pool (the debug
' menu's local STRUCTURE tmp) (S: nothing sent; multistruct_c's takes an aStruct handle)
Public Function SendBuildFinished_struct(psStruct As STRUCTURE) As Boolean
    SendBuildFinished_struct = False
End Function

' --- droid.cpp (droid_c): bool checkValidWeaponForProp(const DROID_TEMPLATE *psTemplate) ---
' (droid.h:281; design.cpp's intCheckValidWeaponForProp calls it.) A DROID_TEMPLATE UDT ByRef, as droid_c's
' droidTemplateType / isTransporter_template take it.
Public Function checkValidWeaponForProp(psTemplate As DROID_TEMPLATE) As Boolean
    Dim psPropStats As Long

    'check propulsion stat for vtol
    psPropStats = DROID_TEMPLATE_getPropulsionStats(psTemplate)

    ' ASSERT_OR_RETURN(false, psPropStats != nullptr, "invalid propulsion stats pointer"): an index is never nullptr

    ' if there are no weapons, then don't even bother continuing
    If psTemplate.numWeaps = 0 Then
        checkValidWeaponForProp = False
        Exit Function
    End If

    If asPropulsionTypes(asPropulsionStats(psPropStats).propulsionType).travel = AIR Then
        'check weapon stat for indirect
        If Not proj_Direct(DROID_TEMPLATE_getWeaponStats(psTemplate, 0)) _
           Or asWeaponStats(DROID_TEMPLATE_getWeaponStats(psTemplate, 0)).vtolAttackRuns = 0 Then
            checkValidWeaponForProp = False
            Exit Function
        End If
    Else
        ' VTOL weapons do not go on non-AIR units.
        If asWeaponStats(DROID_TEMPLATE_getWeaponStats(psTemplate, 0)).vtolAttackRuns <> 0 Then
            checkValidWeaponForProp = False
            Exit Function
        End If
    End If

    'also checks that there is no other system component
    If psTemplate.asParts(COMP_BRAIN) <> 0 _
       And asWeaponStats(DROID_TEMPLATE_getWeaponStats(psTemplate, 0)).weaponSubClass <> WSC_COMMAND Then
        ' assert(false); (NDEBUG in a release build)
        checkValidWeaponForProp = False
        Exit Function
    End If

    checkValidWeaponForProp = True
End Function

' --- display3ddef.h: #define RADTLX (OBJ_BACKX + OBJ_BACKWIDTH + BASE_GAP + 1 +D_W) // Paul's settings (492+12) ---
' (a function: OBJ_BACKX and D_W depend on the screen size; OBJ_BACKX, OBJ_BACKWIDTH, BASE_GAP and D_W are
' hci.h's, fork G1's. mission.cpp and transporter.cpp use it too.)
Public Function RADTLX() As Long
    RADTLX = (OBJ_BACKX + OBJ_BACKWIDTH + BASE_GAP + 1 + D_W)
End Function


' src/mechanics.cpp: void makeAllAvailable() (S: the cheat makes nothing available yet)
Public Sub makeAllAvailable()
End Sub


' lib/netplay: size_t NETgetStatistic(NetStatisticType type, bool sent, bool isTotal = false) (S: 0; WzNet has no statistics)
Public Function NETgetStatistic(ByVal typ As Long, ByVal sent As Boolean, Optional ByVal isTotal As Boolean = False) As Double
    NETgetStatistic = 0
End Function

' src/multimenu.cpp: bool intAddMultiMenu() (S: the in-game multiplayer menu, multimenu.cpp's rest)
Public Function intAddMultiMenu() As Boolean
    intAddMultiMenu = False
End Function


' src/spectatorwidgets.cpp: void specToggleOverlays() (S: the spectator's stats overlay)
Public Sub specToggleOverlays()
End Sub

' src/screens/guidescreen.cpp: bool hasGameGuideTopics() (S: no guide yet)
Public Function hasGameGuideTopics() As Boolean
    hasGameGuideTopics = False
End Function

' src/screens/guidescreen.cpp: bool showGuideScreen(std::function<void ()> onCloseFunc, bool expandedSidebarByDefault) (S)
Public Function showGuideScreen(ByVal onCloseFunc As WzFunc, ByVal expandedSidebarByDefault As Boolean) As Boolean
    showGuideScreen = False
End Function

' src/screens/guidescreen.cpp: void closeGuideScreen() (S)
Public Sub closeGuideScreen()
End Sub

' src/screens/netpregamescreen.cpp: void shutdownGameStartScreen() (S: multiplayer's pre-game screen)
Public Sub shutdownGameStartScreen()
End Sub



' src/activity.cpp: ActivityManager::instance().completedMission(result, collectEndGameStatsData(), cheatsUsed)
' (S: the statistics argument left out, as main_c's ActivityManager_quitGame)
Public Sub ActivityManager_completedMission(ByVal result As Boolean, ByVal cheatsUsed As Boolean)
End Sub




' void renderResearchToBuffer(RESEARCH *psResearch, UDWORD OriginX, UDWORD OriginY, bool rotate = true)
' (psResearch: the research index) (S: nothing drawn)
Public Sub renderResearchToBuffer(ByVal psResearch As Long, ByVal OriginX As Long, ByVal OriginY As Long, Optional ByVal rotate As Boolean = True)
End Sub




' std::shared_ptr<IntFormAnimated> intCreateMultiMenuForm(std::function<void ()> closeButtonHandler = nullptr)
' (S: an empty IntFormAnimated: the intelligence screen of a skirmish attaches it)
Public Function intCreateMultiMenuForm(Optional ByVal closeButtonHandler As WzFunc = Nothing) As WIDGET
    Set intCreateMultiMenuForm = IntFormAnimated_new(True)
End Function


' lib/framework/wzapp.h: WzString wzGetSelection() (the clipboard's text; editbox.cpp's paste).
' (S: the clipboard is not read yet: nothing to paste)
Public Function wzGetSelection() As String
    wzGetSelection = ""
End Function







' lib/netplay/netplay.cpp: int NETshutdown() (S: the net library's NetApi has no NETshutdown yet)
Public Function NETshutdown() As Long
    NETshutdown = 0
End Function

' src/urlhelpers.cpp: bool openURLInBrowser(char const *url) (S)
Public Function openURLInBrowser(ByVal url As String) As Boolean
    openURLInBrowser = False
End Function

' src/urlhelpers.cpp: std::string urlEncode(const char* urlFragment) (S: the fragment unchanged)
Public Function urlEncode(ByVal urlFragment As String) As String
    urlEncode = urlFragment
End Function


' src/notifications.cpp: void addNotification(const WZ_Notification& notification,
' const WZ_Notification_Trigger& trigger); trigger: its timeInterval (0 = Immediate()) (S: cluster G)
Public Sub addNotification(notification As WZ_Notification, ByVal trigger As Long)
End Sub

' src/notifications.cpp: bool hasNotificationsWithTag(const std::string& tag,
' NotificationScope scope = NotificationScope::DISPLAYED_AND_QUEUED) (S: cluster G)
Public Function hasNotificationsWithTag(ByVal tag As String, Optional ByVal scope As NotificationScope = NotificationScope_DISPLAYED_AND_QUEUED) As Boolean
    hasNotificationsWithTag = False
End Function

' src/notifications.cpp: bool cancelOrDismissNotificationIfTag(const std::function<bool (const std::string& tag)>& matchTagFunc,
' NotificationScope scope = NotificationScope::DISPLAYED_AND_QUEUED) (S: cluster G)
Public Function cancelOrDismissNotificationIfTag(ByVal matchTagFunc As WzFunc, Optional ByVal scope As NotificationScope = NotificationScope_DISPLAYED_AND_QUEUED) As Boolean
    cancelOrDismissNotificationIfTag = False
End Function

' lib/framework/wzapp.h: std::vector<unsigned int> wzAvailableDisplayScales(): (VB) into scales(0 .. n - 1),
' n returned (S: WzPlatform has no display scaling yet: 100 only)
Public Function wzAvailableDisplayScales(scales() As Long) As Long
    ReDim scales(0 To 0)
    scales(0) = 100
    wzAvailableDisplayScales = 1
End Function

' lib/framework/wzapp.h: bool wzChangeDisplayScale(unsigned int displayScale) (S: WzPlatform)
Public Function wzChangeDisplayScale(ByVal displayScale As Long) As Boolean
    wzChangeDisplayScale = False
End Function

' lib/framework/wzapp.h: unsigned int wzGetSuggestedDisplayScaleForCurrentWindowSize(unsigned int desiredMaxScreenDimension)
' (S: WzPlatform: the current scale, so the front end never prompts to change it)
Public Function wzGetSuggestedDisplayScaleForCurrentWindowSize(ByVal desiredMaxScreenDimension As Long) As Long
    wzGetSuggestedDisplayScaleForCurrentWindowSize = wzGetCurrentDisplayScale()
End Function


' src/titleui/gamefind.cpp: std::make_shared<WzGameBrowserTitleUI>(parent) (S: later; as above)
Public Function WzGameBrowserTitleUI_new(ByVal parent As WzTitleUI) As WzTitleUI
    Set WzGameBrowserTitleUI_new = WzMsgBoxTitleUI_new("Join Game", "The game browser is not ported yet.", parent)
End Function

' (R) uint8_t NET_numHumanPlayers(void)
Public Function NET_numHumanPlayers() As Long
    Dim Inc As Long, RetVal As Long
    For Inc = 0 To MAX_PLAYERS - 1
        If NetPlay_players_allocated(Inc) And Not NetPlay_players_isSpectator(Inc) Then RetVal = RetVal + 1
    Next
    NET_numHumanPlayers = RetVal
End Function

' (R) std::vector<uint8_t> NET_getHumanPlayers(void): RetVal(0 .. return value - 1)
Public Function NET_getHumanPlayers(RetVal() As Long) As Long
    Dim Inc As Long, n As Long
    Erase RetVal
    For Inc = 0 To MAX_PLAYERS - 1
        If NetPlay_players_allocated(Inc) And Not NetPlay_players_isSpectator(Inc) Then
            ReDim Preserve RetVal(0 To n)
            RetVal(n) = Inc
            n = n + 1
        End If
    Next
    NET_getHumanPlayers = n
End Function

' SpectatorInfo NETGameGetSpectatorInfo() (packed, toUint32)
Public Function NETGameGetSpectatorInfo() As Long
    NETGameGetSpectatorInfo = SpectatorInfo_currentNetPlayState()
End Function

' bool NETcanOpenNewSpectatorSlot()
Public Function NETcanOpenNewSpectatorSlot() As Boolean
    NETcanOpenNewSpectatorSlot = False
End Function

' bool NETmovePlayerToSpectatorOnlySlot(uint32_t playerIdx, bool hostOverride = false)
Public Function NETmovePlayerToSpectatorOnlySlot(ByVal playerIdx As Long, Optional ByVal hostOverride As Boolean = False) As Boolean
    NETmovePlayerToSpectatorOnlySlot = False
End Function

' SpectatorToPlayerMoveResult NETmoveSpectatorToPlayerSlot(uint32_t playerIdx, optional<uint32_t> newPlayerIdx,
' bool hostOverride = false): newPlayerIdx -1 = nullopt
Public Function NETmoveSpectatorToPlayerSlot(ByVal playerIdx As Long, ByVal newPlayerIdx As Long, Optional ByVal hostOverride As Boolean = False) As SpectatorToPlayerMoveResult
    NETmoveSpectatorToPlayerSlot = SpectatorToPlayerMoveResult_FAILED
End Function

' void NETchangePlayerName(UDWORD index, char *newName): the name kept locally (the broadcast: cluster J)
Public Sub NETchangePlayerName(ByVal index As Long, ByVal newName As String)
    setPlayerName index, newName
End Sub

' void NETremRedirects()
Public Sub NETremRedirects()
End Sub

' void NET_clearDownloadingWZFiles()
Public Sub NET_clearDownloadingWZFiles()
End Sub

' NET_getDownloadingWzFiles().empty(): no downloads
Public Function NET_getDownloadingWzFiles_empty() As Boolean
    NET_getDownloadingWzFiles_empty = True
End Function

' PLAYER::resetAll() of NetPlay.players[i]
Public Sub NetPlay_players_resetAll(ByVal i As Long)
End Sub

' NetPlay.scriptSetPlayerDataStrings: size() and [i].clear() (none kept)
Public Function NetPlay_scriptSetPlayerDataStrings_size() As Long
    NetPlay_scriptSetPlayerDataStrings_size = 0
End Function

Public Sub NetPlay_scriptSetPlayerDataStrings_clear(ByVal i As Long)
End Sub

' NetPlay.players[i].wzFiles: erase the files with this hash (NET_FILE_CANCELLED; the map downloads)
Public Sub NetPlay_players_wzFiles_eraseHash(ByVal i As Long, hash As Sha256)
End Sub

' (R) const char *getPlayerColourName(int player)
Public Function getPlayerColourName(ByVal player As Long) As String
    ' STATIC_ASSERT(MAX_PLAYERS <= ARRAY_SIZE(playerColors));
    If Not (player < 16) Then ASSERT_FAILURE "player < ARRAY_SIZE(playerColors)", "src/multiplay.cpp:2730", "getPlayerColourName", "player number (%d) exceeds maximum (%lu)", player, 16
    If player >= 16 Then
        getPlayerColourName = ""
        Exit Function
    End If
    Select Case getPlayerColour(player)
    Case 0: getPlayerColourName = "Green"
    Case 1: getPlayerColourName = "Orange"
    Case 2: getPlayerColourName = "Grey"
    Case 3: getPlayerColourName = "Black"
    Case 4: getPlayerColourName = "Red"
    Case 5: getPlayerColourName = "Blue"
    Case 6: getPlayerColourName = "Pink"
    Case 7: getPlayerColourName = "Cyan"
    Case 8: getPlayerColourName = "Yellow"
    Case 9: getPlayerColourName = "Purple"
    Case 10: getPlayerColourName = "White"
    Case 11: getPlayerColourName = "Bright blue"
    Case 12: getPlayerColourName = "Neon green"
    Case 13: getPlayerColourName = "Infrared"
    Case 14: getPlayerColourName = "Ultraviolet"
    Case 15: getPlayerColourName = "Brown"
    End Select
End Function

' (R) bool isPlayerMuted(uint32_t sender)
Public Function isPlayerMuted(ByVal sender As Long) As Boolean
    If Not ULT(sender, MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS) Then ASSERT_FAILURE "sender < MAX_CONNECTED_PLAYERS", "src/multiplay.cpp:2044", "isPlayerMuted", "Invalid sender: %u", sender: Exit Function
    isPlayerMuted = ingame.muteChat(sender) Or Not ingame.hostChatPermissions(sender)
End Function

' (R, but for storePlayerMuteOption: the known players' database) void setPlayerMuted(uint32_t playerIdx, bool muted)
Public Sub setPlayerMuted(ByVal playerIdx As Long, ByVal muted As Boolean)
    If Not ULT(playerIdx, MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/multiplay.cpp:2024", "setPlayerMuted", "Invalid playerIdx: %u", playerIdx: Exit Sub
    If muted = ingame.muteChat(playerIdx) Then
        ' no change
        Exit Sub
    End If
    ingame.muteChat(playerIdx) = muted
End Sub

' (R) void printConsoleNameChange(const char *oldName, const char *newName)
Public Sub printConsoleNameChange(ByVal oldName As String, ByVal newName As String)
    If game.blindMode <> BLIND_MODE_NONE Then
        Exit Sub
    End If
    displayRoomSystemMessage oldName & " " & ChrW$(&H2192) & " " & newName
End Sub

' bool NetworkTextMessage::receive(NETQUEUE queue): the chat (cluster J)
Public Function NetworkTextMessage_receive(message As NetworkTextMessage, ByVal queue As Long) As Boolean
    NetworkTextMessage_receive = False
End Function

' void MultiPlayerLeave(UDWORD playerIndex) (multijoin.cpp)
Public Sub MultiPlayerLeave(ByVal playerIndex As Long)
End Sub

' void sendLeavingMsg() (multiplay.cpp)
' multivote.cpp (the kick votes, cluster J): void resetKickVoteData()
Public Sub resetKickVoteData()
End Sub

Public Sub sendLeavingMsg()
End Sub

' bool sendMap() (multiplay.cpp: the map downloads)
Public Function sendMap() As Boolean
    sendMap = True
End Function

' bool sendScoreCheck() (multistat.cpp)
Public Function sendScoreCheck() As Boolean
    sendScoreCheck = True
End Function

' bool recvMapFileRequested(NETQUEUE queue), bool recvMapFileData(NETQUEUE queue) (multiplay.cpp)
Public Function recvMapFileRequested(ByVal queue As Long) As Boolean
    recvMapFileRequested = False
End Function

Public Function recvMapFileData(ByVal queue As Long) As Boolean
    recvMapFileData = False
End Function

' bool recvOptions(NETQUEUE queue) (multiopt.cpp: the client's options)
Public Function recvOptions(ByVal queue As Long) As Boolean
    recvOptions = False
End Function

' multivote.cpp: the lobby's votes
Public Function getLobbyChangeVoteTotal() As Long
    getLobbyChangeVoteTotal = 0
End Function

Public Sub startLobbyChangeVote()
End Sub

Public Sub resetLobbyChangePlayerVote(ByVal player As Long)
End Sub

Public Sub sendLobbyChangeVoteData(ByVal lobbyStateChangeVote As Long)
End Sub

Public Sub cancelOrDismissVoteNotifications()
End Sub

' multioptprefs (multiplay): void saveMultiOptionPrefValues(const char *sPlayerName, uint32_t playerIdx)
Public Sub saveMultiOptionPrefValues(ByVal sPlayerName As String, ByVal playerIdx As Long)
End Sub

' multilobbycommands.cpp: bool processChatLobbySlashCommands(const NetworkTextMessage& message,
' HostLobbyOperationsInterface& cmdInterface) (the host's lobby commands: cluster J)
Public Function processChatLobbySlashCommands(message As NetworkTextMessage) As Boolean
    processChatLobbySlashCommands = False
End Function

' main.cpp / stdinreader: void cmdInterfaceLogChatMsg(const NetworkTextMessage& message, const char* log_prefix)
Public Sub cmdInterfaceLogChatMsg(message As NetworkTextMessage, ByVal log_prefix As String)
End Sub

' multiplay.cpp: void autoLobbyNotReadyKickRoutine(std::chrono::steady_clock::time_point now)
Public Sub autoLobbyNotReadyKickRoutine(ByVal now_ As Currency)
End Sub

' banlist: void addIPToBanList(const char *ip, const char *name)
Public Sub addIPToBanList(ByVal ip As String, ByVal name_ As String)
End Sub

' activity.h: ActivityManager::instance().hostKickPlayer(NetPlay.players[player_id], type, reason)
Public Sub ActivityManager_hostKickPlayer(ByVal player As Long, ByVal typ As LOBBY_ERROR_TYPES, ByVal reason As String)
End Sub

' ActivityManager::instance().hostLobbyQuit(errorResult), joinedLobbyQuit(errorResult)
Public Sub ActivityManager_hostLobbyQuit(ByVal errorResult As LOBBY_ERROR_TYPES)
End Sub

Public Sub ActivityManager_joinedLobbyQuit(ByVal errorResult As LOBBY_ERROR_TYPES)
End Sub

' ActivityManager::instance().updateMultiplayGameData(game, ingame, NETGameIsLocked())
Public Sub ActivityManager_updateMultiplayGameData()
End Sub






' src/hci/teamstrategy.cpp (multiplayer: cluster J; quickchat_c calls it):
' int32_t checkedGetPlayerTeam(int32_t i) (REAL: teamstrategy.cpp:128)
Public Function checkedGetPlayerTeam(ByVal i As Long) As Long
    If alliancesSetTeamsBeforeGame(game.alliance) Then
        checkedGetPlayerTeam = NetPlay_players_team(i)
    Else
        checkedGetPlayerTeam = i
    End If
End Function


' std::shared_ptr<WIDGET> makeWzPlayerBlindWaitingRoom(const std::shared_ptr<WzMultiplayerOptionsTitleUI>& titleUI):
' an empty widget
Public Function makeWzPlayerBlindWaitingRoom(ByVal titleUI As WzMultiplayerOptionsTitleUI) As WIDGET
    Set makeWzPlayerBlindWaitingRoom = WIDGET_new()
End Function

' int32_t getWzPlayerBlindWaitingRoomPlayerRowY0(const std::shared_ptr<WIDGET>& blindWaitingRoomWidget)
Public Function getWzPlayerBlindWaitingRoomPlayerRowY0(ByVal blindWaitingRoomWidget As WIDGET) As Long
    getWzPlayerBlindWaitingRoomPlayerRowY0 = 0
End Function

' lib/framework/wzapp.h: bool wzSeemsLikeNonTouchPlatform()
Public Function wzSeemsLikeNonTouchPlatform() As Boolean
    wzSeemsLikeNonTouchPlatform = True
End Function

' (R) lib/framework/string_ext.h: size_t nthOccurrenceOfChar(const std::string& str, const char c, int n): the
' 0-based position, -1 (std::string::npos) when there are fewer
Public Function nthOccurrenceOfChar(ByVal str_ As String, ByVal c As String, ByVal n As Long) As Long
    Dim i As Long, count As Long
    For i = 1 To Len(str_)
        If Mid$(str_, i, 1) = c Then
            count = count + 1
            If count = n Then
                nthOccurrenceOfChar = i - 1
                Exit Function
            End If
        End If
    Next
    nthOccurrenceOfChar = -1
End Function



' src/urlhelpers.cpp: bool openFolderInDefaultFileManager(const char* path) (Windows: ShellExecuteW "open")
Public Function openFolderInDefaultFileManager(ByVal path As String) As Boolean
    openFolderInDefaultFileManager = False
End Function



' std::vector<video_backend> wzAvailableGfxBackends(): the count, the backends in result(0 .. count - 1)
Public Function wzAvailableGfxBackends(result() As Long) As Long
    ReDim result(0 To 0)
    result(0) = video_backend_opengl
    wzAvailableGfxBackends = 1
End Function

' void wzPostChangedSwapInterval()
Public Sub wzPostChangedSwapInterval()
End Sub

' bool wzIsSupportedWindowMode(WINDOW_MODE mode)
Public Function wzIsSupportedWindowMode(ByVal mode As Long) As Boolean
    wzIsSupportedWindowMode = True
End Function

' void wzGetWindowResolution(int *screen, unsigned int *width, unsigned int *height)
Public Sub wzGetWindowResolution(screen As Long, width As Long, height As Long)
    screen = 0
    width = screenWidth
    height = screenHeight
End Sub

' std::vector<optional<screeninfo>> wzAvailableResolutions(): the count, modes(i) with has_(i) (False: nullopt)
Public Function wzAvailableResolutions(modes() As screeninfo, has_() As Boolean) As Long
    ReDim modes(0 To 0)
    ReDim has_(0 To 0)
    has_(0) = False                         ' nullopt: "Desktop Full"
    wzAvailableResolutions = 1
End Function

' optional<screeninfo> wzGetCurrentFullscreenDisplayMode(): has_value, the value in info
Public Function wzGetCurrentFullscreenDisplayMode(info As screeninfo) As Boolean
    wzGetCurrentFullscreenDisplayMode = False
End Function

' bool wzChangeFullscreenDisplayMode(optional<screeninfo> config)
Public Function wzChangeFullscreenDisplayMode(ByVal has_config As Boolean, config As screeninfo) As Boolean
    wzChangeFullscreenDisplayMode = True
End Function

' MinimizeOnFocusLossBehavior wzGetCurrentMinimizeOnFocusLossBehavior()
Public Function wzGetCurrentMinimizeOnFocusLossBehavior() As Long
    wzGetCurrentMinimizeOnFocusLossBehavior = war_getMinimizeOnFocusLoss()
End Function

' void wzSetMinimizeOnFocusLoss(MinimizeOnFocusLossBehavior behavior)
Public Sub wzSetMinimizeOnFocusLoss(ByVal behavior As Long)
End Sub

' unsigned int wzGetMaximumDisplayScaleForCurrentWindowSize()
Public Function wzGetMaximumDisplayScaleForCurrentWindowSize() As Long
    wzGetMaximumDisplayScaleForCurrentWindowSize = 500
End Function

' bool wzChangeCursorScale(unsigned int cursorScale) (stand-in: only the current scale "succeeds")
Public Function wzChangeCursorScale(ByVal cursorScale As Long) As Boolean
    wzChangeCursorScale = (cursorScale = war_getCursorScale())
End Function




























' bool shouldSkipReadyResetOnPlayerJoinLeaveEvent()
Public Function shouldSkipReadyResetOnPlayerJoinLeaveEvent() As Boolean
    Dim mode As Long
    ' If a player joins or leaves, do not reset existing player "ready" status if:
    ' 1. It's a blind simple lobby (i.e. waiting room)
    ' 2. min_autostart_player_count is set (ex. via the --startplayers= command line option) *and* locked.readybeforefull is false
    mode = game.blindMode   ' isBlindSimpleLobby(game.blindMode) (multiplaydefs.h)
    shouldSkipReadyResetOnPlayerJoinLeaveEvent = (mode = BLIND_MODE_BLIND_LOBBY_SIMPLE_LOBBY Or mode = BLIND_MODE_BLIND_GAME_SIMPLE_LOBBY) _
        Or ((min_autostart_player_count() > 0) And Not getLockedOptions().readybeforefull)
End Function

' bool isLocallyKnownPlayer(std::string const &name, EcKey const &key): (stand-in) knownPlayersDB is not
' ported, so no player is known
Public Function isLocallyKnownPlayer(ByVal name_ As String, key As EcKey) As Boolean
    isLocallyKnownPlayer = False
End Function

' const EcKey& getOutputPlayerIdentity(UDWORD player): (stand-in) getTruePlayerIdentity(player).identity is
' the multistats identity outside blind mode (the only case the lobby widgets ask for it)
Public Function getOutputPlayerIdentity(ByVal player As Long) As EcKey
    getOutputPlayerIdentity = getMultiStats(player).identity
End Function

' unsigned NETgetDownloadProgress(unsigned player): Returns 100 when done.
Public Function NETgetDownloadProgress(ByVal player As Long) As Long
    NETgetDownloadProgress = 100
End Function

' PLAYER::fileSendInProgress(): wzFiles not empty
Public Property Get NetPlay_players_fileSendInProgress(ByVal i As Long) As Boolean
    NetPlay_players_fileSendInProgress = False
End Property

' bool NETopenNewSpectatorSlot()
Public Function NETopenNewSpectatorSlot() As Boolean
    NETopenNewSpectatorSlot = False
End Function

' void NETresetGamePassword(): (stand-in) NetPlay.gamePassword is netplay_c's Private; only NETGameLocked(false)
Public Sub NETresetGamePassword()
    NETGameLocked False
End Sub

' NetPlay.gamePassword (char[password_string_size]): (stand-in) not exposed by WzNet
Public Function NetPlay_gamePassword() As String
    NetPlay_gamePassword = ""
End Function

' NetPlay.GamePassworded
Public Function NetPlay_GamePassworded() As Boolean
    NetPlay_GamePassworded = NETGameIsLocked()
End Function

' void NETsetLobbyOptField(const char *Value, const NET_LOBBY_OPT_FIELD Field)
Public Sub NETsetLobbyOptField(ByVal value As String, ByVal field As NET_LOBBY_OPT_FIELD)
End Sub

' void NETsetLobbyConfigFlagsFields(uint8_t alliancesType, uint8_t techLevel, uint8_t powerLevel, uint8_t basesLevel)
Public Sub NETsetLobbyConfigFlagsFields(ByVal alliancesType As Long, ByVal techLevel As Long, ByVal powerLevel As Long, ByVal basesLevel As Long)
End Sub

' std::string NET_getCurrentHostedLobbyGameId(): "" when not listed in a lobby
Public Function NET_getCurrentHostedLobbyGameId() As String
    NET_getCurrentHostedLobbyGameId = ""
End Function

' optional<bool> NET_getHostJoinOptionSupportedByLobby(netlobby::HostJoinOptionType optType): has_value, the value ByRef
Public Function NET_getHostJoinOptionSupportedByLobby(ByVal optType As HostJoinOptionType, result As Boolean) As Boolean
    NET_getHostJoinOptionSupportedByLobby = False   ' nullopt
End Function



' LobbyHostDirectJoinOption NETgetLobbyHostDirectJoinOption()
Public Function NETgetLobbyHostDirectJoinOption() As LobbyHostDirectJoinOption
    If Not lobbyDirectJoinOption_set Then
        lobbyDirectJoinOption = LobbyHostDirectJoinOption_Local     ' static ... lobbyDirectJoinOption = LobbyHostDirectJoinOption::Local;
        lobbyDirectJoinOption_set = True
    End If
    NETgetLobbyHostDirectJoinOption = lobbyDirectJoinOption
End Function

' void NETsetLobbyHostDirectJoinOption(LobbyHostDirectJoinOption option)
Public Sub NETsetLobbyHostDirectJoinOption(ByVal opt As LobbyHostDirectJoinOption)
    lobbyDirectJoinOption = opt
    lobbyDirectJoinOption_set = True
End Sub

' void wzSetClipboardText(const char *src)
Public Sub wzSetClipboardText(ByVal src As String)
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText src
    Err.Clear
End Sub

' bool getMultiOptionPrefValue(ScavType / AllianceType / PowerSetting / CampType / TechLevel val)
Public Function getMultiOptionPrefValue_ScavType(ByVal val_ As Long) As Boolean
End Function

Public Function getMultiOptionPrefValue_AllianceType(ByVal val_ As Long) As Boolean
End Function

Public Function getMultiOptionPrefValue_PowerSetting(ByVal val_ As Long) As Boolean
End Function

Public Function getMultiOptionPrefValue_CampType(ByVal val_ As Long) As Boolean
End Function

Public Function getMultiOptionPrefValue_TechLevel(ByVal val_ As Long) As Boolean
End Function

' void setMultiOptionPrefValue(ScavType / ... val, bool pref)
Public Sub setMultiOptionPrefValue_ScavType(ByVal val_ As Long, ByVal pref As Boolean)
End Sub

Public Sub setMultiOptionPrefValue_AllianceType(ByVal val_ As Long, ByVal pref As Boolean)
End Sub

Public Sub setMultiOptionPrefValue_PowerSetting(ByVal val_ As Long, ByVal pref As Boolean)
End Sub

Public Sub setMultiOptionPrefValue_CampType(ByVal val_ As Long, ByVal pref As Boolean)
End Sub

Public Sub setMultiOptionPrefValue_TechLevel(ByVal val_ As Long, ByVal pref As Boolean)
End Sub

' size_t getMultiOptionPrefValueTotal(ScavType / ... val)
Public Function getMultiOptionPrefValueTotal_ScavType(ByVal val_ As Long) As Long
End Function

Public Function getMultiOptionPrefValueTotal_AllianceType(ByVal val_ As Long) As Long
End Function

Public Function getMultiOptionPrefValueTotal_PowerSetting(ByVal val_ As Long) As Long
End Function

Public Function getMultiOptionPrefValueTotal_CampType(ByVal val_ As Long) As Long
End Function

Public Function getMultiOptionPrefValueTotal_TechLevel(ByVal val_ As Long) As Long
End Function

' void setMultiOptionBuiltinPrefValues(bool pref)
Public Sub setMultiOptionBuiltinPrefValues(ByVal pref As Boolean)
End Sub

' void resetMultiOptionPrefValues(uint32_t player)
Public Sub resetMultiOptionPrefValues(ByVal player As Long)
End Sub

' void printStructureLimitsInfo(std::vector<MULTISTRUCTLIMITS>& structureLimits, const PrintStructureLimitsLineFunc& printLineFunc)
' (structureLimits(0 .. structureLimitsCount - 1); printLineFunc: a WzFunc taking the line). (stand-in) only the
' "no limits" case and the !bLimiterLoaded case are as the C++; changed limits print their count.
Public Sub printStructureLimitsInfo(structureLimits() As MULTISTRUCTLIMITS, ByVal structureLimitsCount As Long, ByVal printLineFunc As WzFunc)
    If structureLimitsCount = 0 Then
        WzFunc_call printLineFunc, "Limits set to default."
    ElseIf Not bLimiterLoaded Then
        WzFunc_call printLineFunc, "Host initialized " & CStr(structureLimitsCount) & " limits, unable to show them due to mods"
    Else
        WzFunc_call printLineFunc, "Changed structure limits [" & CStr(structureLimitsCount) & "]:"
    End If
End Sub

' std::make_shared<WzMultiLimitTitleUI>(parent): (stand-in) the parent itself, so the limits button leaves
' the lobby as it is
Public Function WzMultiLimitTitleUI_new(ByVal parent As WzMultiplayerOptionsTitleUI) As WzTitleUI
    Set WzMultiLimitTitleUI_new = parent
End Function
