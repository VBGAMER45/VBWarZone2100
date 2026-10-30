Attribute VB_Name = "multiint_lobby_c"
' src/multiint.cpp (part: the lobby screen): the widget ids of multiint.h, the
' choosers' free functions, the requests of the lobby (team, colour, faction,
' position, slot type, ready) and their receivers, the room's chat box
' (RoomMessage, ChatBoxButton, ChatBoxWidget, ChatBoxPlayerNameWidget,
' addChatBox), the display functions (displayAi, displayDifficulty,
' intDisplayFeBox, ...), WzMultiButton and its makers (addMultiBut,
' makeMultiBut, updateMultiBut), the position choosers (the row factories,
' WzPlayerSelectionChangePositionRow, WzPlayerIndexSwapPositionRow,
' WzPositionChooser), the randomize and map settings helpers, the kicks,
' the ready checks and the room messages. The game state and the start path
' are multiint_c's; the screen's methods WzMultiplayerOptionsTitleUI.cls's.
' The lambdas are WzFuncs run by multiint_lobby_c_fn (FN_ ids 5000-5999); a
' lambda's std::weak_ptr<WzMultiplayerOptionsTitleUI> is the screen's address,
' titleUI_lock gives it back while it is the current lobby screen.
' (VB) widget class ChatBoxButton : W_BUTTON
' (VB) widget class ChatBoxWidget : IntFormAnimated
' (VB) virtual ChatBoxWidget.initialize
' (VB) widget class ChatBoxPlayerNameWidget : WIDGET
' (VB) widget class WzMultiButton : W_BUTTON
' (VB) widget class WzPlayerSelectionChangePositionRow : W_BUTTON
' (VB) widget class WzPlayerIndexSwapPositionRow : W_BUTTON
' (VB) widget class WzPositionChooser : WIDGET
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' lib/netplay/nettypes.h
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' multiint.cpp
Private Const PLAYERBOX_X0 As Long = 7
Private Const KICK_REASON_TAG As String = "kickReason"
Private Const SLOTTYPE_REQUEST_TAG As String = "slotType::request"

' multiint.h
Public Const AI_NOT_FOUND As Long = -99
' (DEFAULT_SKIRMISH_AI_SCRIPT_NAME: stubs_game's)

Public Const CON_TYPESID_START As Long = 10105
Public Const CON_TYPESID_END As Long = 10128
Public Const CON_SETTINGS As Long = 10130
Public Const CON_SETTINGS_LABEL As Long = 10131
Public Const CON_SETTINGSWIDTH As Long = 200
Public Const CON_SETTINGSHEIGHT As Long = 100
Public Const CON_OK As Long = 10101
Public Const CON_CANCEL As Long = 10102
Public Const CON_PHONE As Long = 10132
Public Const CON_PHONEX As Long = 20
Public Const CON_IP As Long = 10133
Public Const CON_IPX As Long = 20
Public Const CON_IPY As Long = 45
Public Const CON_IP_CANCEL As Long = 10134
Public Const CON_SPECTATOR_BOX As Long = 10135
Public Const CON_PASSWORD As Long = 10139
Public Const CON_PASSWORDYES As Long = 10141
Public Const CON_PASSWORDNO As Long = 10142

Public Const GAMES_GAMEHEADER As Long = 10200
Public Const GAMES_GAMELIST As Long = 10201
Public Const GAMES_MAX As Long = 100
Public Const GAMES_GAMEWIDTH As Long = 525
Public Const GAMES_GAMEHEIGHT As Long = 28
Public Const GAMES_STATUS_START As Long = 378
Public Const GAMES_GAMENAME_START As Long = 2
Public Const GAMES_MAPNAME_START As Long = 168
Public Const GAMES_MODNAME_START As Long = 168 + 6     ' indent a bit
Public Const GAMES_PLAYERS_START As Long = 342

Public Const MULTIOP_PLAYERS As Long = 10231
Public Const MULTIOP_PLAYERSX As Long = 315
Public Const MULTIOP_PLAYERSY As Long = 1
Public Const MULTIOP_PLAYER_START As Long = 102350&    'list of players
Public Const MULTIOP_PLAYER_END As Long = 102381&
Public Const MULTIOP_PLAYERSW As Long = 298
Public Const MULTIOP_PLAYERS_TABS As Long = 10232
Public Const MULTIOP_PLAYERS_TABS_H As Long = 24
Public Const MULTIOP_PLAYERSH As Long = (384 + MULTIOP_PLAYERS_TABS_H + 1)
Public Const MULTIOP_BLIND_WAITING_ROOM As Long = 10233

Public Const MULTIOP_ROW_WIDTH As Long = 298

Public Const MULTIOP_TEAMS_START As Long = 102310&     'List of teams
Public Const MULTIOP_TEAMS_END As Long = 102341&
Public Const MULTIOP_TEAMSWIDTH As Long = 28
Public Const MULTIOP_TEAMSHEIGHT As Long = 38

Public Const MULTIOP_TEAMCHOOSER_FORM As Long = 102800&
Public Const MULTIOP_TEAMCHOOSER As Long = 102810&
Public Const MULTIOP_TEAMCHOOSER_END As Long = 102841&
Public Const MULTIOP_TEAMCHOOSER_KICK As Long = 10289
Public Const MULTIOP_TEAMCHOOSER_SPECTATOR As Long = 10288
Public Const MULTIOP_TEAMCHOOSER_BAN As Long = 10287

Public Const MULTIOP_INLINE_OVERLAY_ROOT_FRM As Long = 10286

Public Const MULTIOP_READY_FORM_ID As Long = 102900&
Public Const MULTIOP_READY_START As Long = (MULTIOP_READY_FORM_ID + MAX_CONNECTED_PLAYERS + 1)
Public Const MULTIOP_READY_END As Long = (MULTIOP_READY_START + MAX_CONNECTED_PLAYERS - 1)
Public Const MULTIOP_READY_WIDTH As Long = 41
Public Const MULTIOP_READY_HEIGHT As Long = 38

Public Const MULTIOP_PLAYERWIDTH As Long = 282
Public Const MULTIOP_PLAYERHEIGHT As Long = 38

Public Const MULTIOP_OPTIONS As Long = 10250
Public Const MULTIOP_OPTIONSX As Long = 32
Public Const MULTIOP_OPTIONSY As Long = 1
Public Const MULTIOP_OPTIONSW As Long = 284
Public Const MULTIOP_OPTIONSH As Long = MULTIOP_PLAYERSH

Public Const MULTIOP_EDITBOXW As Long = 201
Public Const MULTIOP_EDITBOXH As Long = 30
Public Const MULTIOP_SEARCHBOXH As Long = 15

Public Const MULTIOP_BLUEFORMW As Long = 231

Public Const MROW1 As Long = 6
Public Const MCOL0 As Long = 45
Public Const MCOL1 As Long = (MCOL0 + 26 + 10)          ' rem 10 for 4 lines.
Public Const MCOL2 As Long = (MCOL1 + 38)
Public Const MCOL3 As Long = (MCOL2 + 38)
Public Const MCOL4 As Long = (MCOL3 + 38)

Public Const MULTIOP_PNAME_ICON As Long = 10252
Public Const MULTIOP_PNAME As Long = 10253
Public Const MULTIOP_MAP As Long = 10259
Public Const MULTIOP_STRUCTLIMITS As Long = 21277       ' we are using 10277 already

Public Const MULTIOP_CANCELX As Long = 6
Public Const MULTIOP_CANCELY As Long = 6

Public Const MULTIOP_CHATBOX As Long = 10278
Public Const MULTIOP_CHATBOXX As Long = MULTIOP_OPTIONSX
Public Const MULTIOP_CHATBOXY As Long = MULTIOP_PLAYERSH
Public Const MULTIOP_CHATBOXW As Long = ((MULTIOP_PLAYERSX + MULTIOP_PLAYERSW) - MULTIOP_OPTIONSX)

Public Const MULTIOP_CONSOLEBOX As Long = &H1A001      ' TODO: these should be enums!
Public Const MULTIOP_CONSOLEBOXX As Long = MULTIOP_OPTIONSX
Public Const MULTIOP_CONSOLEBOXY As Long = 422
Public Const MULTIOP_CONSOLEBOXW As Long = ((MULTIOP_PLAYERSX + MULTIOP_PLAYERSW) - MULTIOP_OPTIONSX)
Public Const MULTIOP_CONSOLEBOXH As Long = 64

Public Const MULTIOP_CHATEDIT As Long = 10279
Public Const MULTIOP_CHATEDITX As Long = 4
Public Const MULTIOP_CHATEDITH As Long = 20
Public Const MULTIOP_CHATEDITW As Long = (MULTIOP_CHATBOXW - 8)

Public Const MULTIOP_COLCHOOSER_FORM As Long = 10280
Public Const MULTIOP_COLCHOOSER As Long = 102711&       '10281
Public Const MULTIOP_COLCHOOSER_END As Long = 102742&   '10288

Public Const MULTIOP_COLOUR_START As Long = 10332
Public Const MULTIOP_COLOUR_END As Long = (MULTIOP_COLOUR_START + MAX_PLAYERS)
Public Const MULTIOP_COLOUR_WIDTH As Long = 29

Public Const MULTIOP_AI_FORM As Long = (MULTIOP_COLOUR_END + 1)
Public Const MULTIOP_AI_START As Long = (MULTIOP_AI_FORM + 1)
Public Const MULTIOP_AI_END As Long = (MULTIOP_AI_START * MAX_PLAYERS)
Public Const MULTIOP_AI_OPEN As Long = (MULTIOP_AI_END + 1)
Public Const MULTIOP_AI_CLOSED As Long = (MULTIOP_AI_END + 2)
Public Const MULTIOP_AI_SPECTATOR As Long = (MULTIOP_AI_END + 3)

Public Const MULTIOP_DIFFICULTY_INIT_START As Long = (MULTIOP_AI_END + 4)
Public Const MULTIOP_DIFFICULTY_INIT_END As Long = (MULTIOP_DIFFICULTY_INIT_START + MAX_PLAYERS)

Public Const MULTIOP_DIFFICULTY_CHOOSE_START As Long = (MULTIOP_DIFFICULTY_INIT_END + 1)
Public Const MULTIOP_DIFFICULTY_CHOOSE_END As Long = (MULTIOP_DIFFICULTY_CHOOSE_START + MAX_PLAYERS)

Public Const MULTIOP_ADD_SPECTATOR_SLOTS As Long = (MULTIOP_DIFFICULTY_CHOOSE_END + 1)

Public Const MULTIOP_FACTION_START As Long = (MULTIOP_ADD_SPECTATOR_SLOTS + 100000)
Public Const MULTIOP_FACTION_END As Long = (MULTIOP_FACTION_START + MAX_PLAYERS)
Public Const MULTIOP_FACTION_WIDTH As Long = 31

Public Const MULTIOP_FACCHOOSER As Long = (MULTIOP_FACTION_END + 1)
Public Const MULTIOP_FACCHOOSER_END As Long = (MULTIOP_FACCHOOSER + NUM_FACTIONS)
Public Const MULTIOP_FACCHOOSER_FORM As Long = (MULTIOP_FACCHOOSER_END + 1)

Public Const CON_NAMEBOXWIDTH As Long = CON_SETTINGSWIDTH - CON_PHONEX
Public Const CON_NAMEBOXHEIGHT As Long = 15

Public Const MULTIOP_OKW As Long = 37
Public Const MULTIOP_OKH As Long = 24

Public Const MULTIOP_BUTW As Long = 35
Public Const MULTIOP_BUTH As Long = 24

Public Const CON_OKX As Long = CON_SETTINGSWIDTH - MULTIOP_OKW * 2 - 13
Public Const CON_OKY As Long = CON_SETTINGSHEIGHT - MULTIOP_OKH - 3

' enum RoomMessageType
Public Enum RoomMessageType
    RoomMessagePlayer
    RoomMessageSystem
    RoomMessageNotify
End Enum

' struct RoomMessage (RoomMessage::player(sender, text), ::system(text), ::notify(text): RoomMessage_player,
' RoomMessage_system, RoomMessage_notify). sender: the PlayerReference's player index (NetPlay.playerReferences:
' the identities come with cluster J; -1 = nullptr)
Public Type RoomMessage
    sender As Long
    text As String
    time_ As Double                     ' std::time_t time = std::time(nullptr);
    typ As RoomMessageType
End Type

' ChatBoxWidget::ChatBoxSendMode
Public Enum ChatBoxSendMode
    ChatBoxSendMode_DISABLED
    ChatBoxSendMode_HOSTMSG_ONLY
    ChatBoxSendMode_ENABLED
End Enum

' (VB) the WzPlayerSelectPositionRowFactory subclasses (their virtual make: a Select Case)
Public Enum WzPlayerSelectPositionRowFactory
    WzPlayerSelectionChangePositionRowFactory
    WzPlayerIndexSwapPositionRowRactory
End Enum

' std::vector<RoomMessage> ChatBoxWidget::persistentMessageLocalStorage;
Private persistentMessageLocalStorage() As RoomMessage
Private persistentMessageLocalStorage_size As Long

' the FN ids of this module's lambdas and function pointers (WzFunc): args(0) is the widget (W_BUTTON&,
' WIDGET *psWidget, W_EDITBOX&, Paragraph&) where the C++ passes one; a display function's args are
' (psWidget, xOffset, yOffset)
Public Const FN_addInlineChooserBlueForm_calcLayout As Long = 5000 ' [x, y, psWeakParent](WIDGET *psWidget)
Public Const FN_intDisplayFeBox As Long = 5001
Public Const FN_intDisplayFeBox_Spectator As Long = 5002
Public Const FN_intDisplayFeBox_SpectatorOnly As Long = 5003
Public Const FN_displayAi As Long = 5004
Public Const FN_displayDifficulty As Long = 5005
Public Const FN_displayChatEdit As Long = 5006
Public Const FN_displaySpectatorAddButton As Long = 5007
Public Const FN_displayMultiEditBox As Long = 5008
Public Const FN_startHosting_onClick As Long = 5010         ' addGameOptions: [psWeakTitleUI](W_BUTTON&)
Public Const FN_multiLobbyOptionsForm_calcLayout As Long = 5011 ' [psWeakTitleUI](WIDGET *psWidget)
Public Const FN_optionsForm_calcLayout As Long = 5012       ' LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_rightSideChooser_calcLayout As Long = 5013  ' [psWeakParent](WIDGET *psWidget) (cap(1): the x adjustment, 0 or -3)
Public Const FN_difficultyRow_onClick As Long = 5014        ' openDifficultyChooser: [psWeakTitleUI, difficultyIdx, player]
Public Const FN_task_closeDifficultyChooser As Long = 5015  ' widgScheduleTask([pStrongPtr] {...})
Public Const FN_aiOpenClose_onClick As Long = 5016          ' openAiChooser: [psWeakTitleUI, player](W_BUTTON& clickedButton)
Public Const FN_task_closeAiChooser As Long = 5017
Public Const FN_aiList_calcLayout As Long = 5018            ' [aiListStartXPos, aiListStartYPos, aiListEntryWidth, aiListHeight]
Public Const FN_aiRow_onClick As Long = 5019                ' [psWeakTitleUI, aiIdx, player]
Public Const FN_initPUserData_DisplayAICache As Long = 5020 ' []() -> void * { return new DisplayAICache(); }
Public Const FN_changePositionRow_onClick As Long = 5021    ' WzPlayerSelectionChangePositionRow::make: [switcherPlayerIdx, titleUI]
Public Const FN_task_closePositionChooser As Long = 5022
Public Const FN_swapPositionRow_onClick As Long = 5023      ' WzPlayerIndexSwapPositionRow::make: [switcherPlayerIdx, titleUI]
Public Const FN_task_closePlayerSlotSwapChooser As Long = 5024
Public Const FN_positionChooser_clicked_task As Long = 5025 ' WzPositionChooser::clicked: [psWeakTitleUICopy]
Public Const FN_team_onClick As Long = 5026                 ' openTeamChooser: [player, psWeakTitleUI](W_BUTTON &button)
Public Const FN_task_closeTeamChooser_addPlayerBox As Long = 5027
Public Const FN_kick_onClick As Long = 5028
Public Const FN_task_closeTeamChooser As Long = 5029
Public Const FN_ban_onClick As Long = 5030
Public Const FN_specHost_onClick As Long = 5031
Public Const FN_specClient_onClick As Long = 5032
Public Const FN_askToPlay_onClick As Long = 5033            ' (cap(2): the C++ line of the host's or the client's lambda)
Public Const FN_colour_onClick As Long = 5034               ' openColourChooser: [player, psWeakTitleUI](W_BUTTON &button)
Public Const FN_task_closeColourChooser As Long = 5035
Public Const FN_inlineOverlay_onClicked As Long = 5036      ' start: newRootFrm->onClickedFunc = [psWeakTitleUI]() (and onCancelPressed)
Public Const FN_inlineOverlay_task As Long = 5037           ' its widgScheduleTask([psWeakTitleUI]{...})
Public Const FN_playerSlotSwapChooser_calcLayout As Long = 5039 ' openPlayerSlotSwapChooser's calcLayouts (cap(0): which)
Public Const FN_addSpectatorSlots_onClick As Long = 5040    ' addPlayerBox: [psWeakTitleUI](W_BUTTON& clickedButton)
Public Const FN_playersForm_calcLayout As Long = 5041       ' LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_playersTabs_calcLayout As Long = 5042       ' LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_blindWaitingRoom_calcLayout As Long = 5043  ' LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_chatBox_calcLayout As Long = 5044           ' addChatBox: LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_chatBox_onEditReturn As Long = 5045         ' ChatBoxWidget::initialize: [weakSelf](W_EDITBOX& widg)
Public Const FN_chatBox_quickChat_onClick As Long = 5046    ' [weakSelf](W_BUTTON& widg)
Public Const FN_chatBox_options_onClick As Long = 5047      ' [weakSelf](W_BUTTON& widg)
Public Const FN_chatBox_consoleListener As Long = 5048      ' [&](ConsoleMessage const &message): args (sender, text)
Public Const FN_chatBox_onTab As Long = 5049                ' [weakSelf](W_EDITBOX& widg) -> bool
Public Const FN_chatBox_paragraphContextMenu As Long = 5050 ' displayMessage: [msgText, senderCopy](Paragraph& paragraph, WIDGET_KEY key)
Public Const FN_addSpectatorSlots_task As Long = 5051       ' widgScheduleTask([psWeakTitleUI]{...})
Public Const FN_chatBox_openQuickChat_task As Long = 5052   ' widgScheduleTask([weakSelf]() {... openQuickChatOverlay(); }) (cap(1): the C++ line)
Public Const FN_chatBox_displayOptions_task As Long = 5053  ' widgScheduleTask([weakSelf]() {... displayOptionsOverlay(...); })
Public Const FN_chatBox_overlay_onClicked As Long = 5054    ' newRootFrm->onClickedFunc = [psWeakChatBoxWidget]() (and onCancelPressed)
Public Const FN_chatBox_closeQuickChat_task As Long = 5055  ' widgScheduleTask([psWeakChatBoxWidget]{...}) (cap(1): the C++ line)
Public Const FN_backdrop_calcLayout As Long = 5056          ' calcBackdropLayoutForMultiplayerOptionsTitleUI
Public Const FN_faction_onClick As Long = 5057              ' openColourChooser: [player, psWeakTitleUI](W_BUTTON &button)
Public Const FN_chatBox_quickChatSent As Long = 5058        ' createQuickChatForm's [psWeakChatBoxWidget]() (on quick chat sent)
Public Const FN_chatBox_quickChatForm_calcLayout As Long = 5059 ' [psWeakChatBoxWidget](WIDGET* psWidget)
Public Const FN_chatBox_optionsOverlay_onClicked As Long = 5060 ' [psWeakOptionsOverlayScreen, psWeakChatBoxWidget]()
Public Const FN_chatBox_chatOptionsChanged As Long = 5061   ' createChatOptionsForm's []() { addChatBox(); }
Public Const FN_chatBox_chatOptionsForm_calcLayout As Long = 5062 ' [psWeakChatBoxWidget](WIDGET* psWidget)
Public Const FN_chatBox_menuButton_onClick As Long = 5063   ' createOptionsButton: [weakChatBoxWidget, onClickFunc](W_BUTTON& button)
Public Const FN_chatBox_copyText As Long = 5064             ' [weakChatBoxWidget, textToCopy](W_BUTTON& button)
Public Const FN_chatBox_mutePlayer As Long = 5065           ' [sender, currentlyMuted](W_BUTTON& button)

' openPlayerSlotSwapChooser's calcLayouts (FN_playerSlotSwapChooser_calcLayout's cap(0))
Private Const SWAPLAYOUT_AIFORM As Long = 0
Private Const SWAPLAYOUT_POSITIONCHOOSER As Long = 1
Private Const SWAPLAYOUT_CONTEXTFORM As Long = 2
Private Const SWAPLAYOUT_PLAYERROW As Long = 3
Private Const SWAPLAYOUT_ROWLABEL As Long = 4
Private Const SWAPLAYOUT_SWITCHARROW As Long = 5
Private Const SWAPLAYOUT_SIDETEXT As Long = 6

' enum class SwapPlayerIndexesResult
Public Enum SwapPlayerIndexesResult
    SwapPlayerIndexesResult_SUCCESS
    SwapPlayerIndexesResult_FAILURE
    SwapPlayerIndexesResult_ERROR_HOST_MADE_US_PLAYER_WITHOUT_PERMISSION
End Enum

' constexpr int MULTIOP_CHATEDITAREA_*
Private Const MULTIOP_CHATEDITAREA_INSET As Long = 5
Private Const MULTIOP_CHATEDITAREA_INTERNAL_PADDING As Long = 6
Private Const MULTIOP_CHATEDITAREA_HEIGHT As Long = MULTIOP_CHATEDITH + (MULTIOP_CHATEDITAREA_INSET * 2)

Private Const MENU_BUTTONS_PADDING As Long = 20

' multiplay.h, ai.h: the options' constexpr maxima (ScavType, AllianceType, PowerSetting, CampType, TechLevel)
Private Const SCAV_TYPE_MAX As Long = ULTIMATE_SCAVENGERS
Private Const ALLIANCE_TYPE_MAX As Long = ALLIANCES_UNSHARED
Private Const POWER_SETTING_MAX As Long = LEV_HI
Private Const CAMP_TYPE_MAX As Long = CAMP_WALLS
Private Const TECH_LEVEL_MIN As Long = 1                    ' TechLevel::TECH_1
Private Const TECH_LEVEL_MAX As Long = 4                    ' TechLevel::TECH_4

' multilobbycommands.h
Private Const LOBBY_COMMAND_PREFIX As String = "/"
Private Const LOBBY_COMMAND_PREFIX_LENGTH As Long = 1

' multiint.cpp's SLOTTYPE_TAG_PREFIX (multiint_c's is Private)

' WzMultiplayerOptionsTitleUI::MultiMessagesResult
Public Enum MultiMessagesResult
    MultiMessagesResult_Continue
    MultiMessagesResult_StoppedJoining
End Enum

Public Const MAP_PREVIEW_DISPLAY_TIME As Long = 2500   ' number of milliseconds to show map in preview

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private perfFrequency As Currency       ' (VB) QueryPerformanceFrequency: steady_clock durations in seconds

'--------------------------------------------------------------------------
' (VB) the lambdas' std::weak_ptr<WzMultiplayerOptionsTitleUI>: the screen if p is still the current lobby
' screen (multiint_c's currentMultiOptionsTitleUI), else Nothing (lock() of an expired pointer)
Public Function titleUI_lock(ByVal p As Long) As WzMultiplayerOptionsTitleUI
    Dim o As WzMultiplayerOptionsTitleUI
    If p = 0 Or p <> currentMultiOptionsTitleUI Then Exit Function
    CopyMemory o, p, 4
    Set titleUI_lock = o
    CopyMemory o, 0&, 4
End Function

' (VB) lib/netplay/netplay.h's SpectatorInfo, packed as its toUint32(): availableSpectatorSlots()
Public Function SpectatorInfo_availableSpectatorSlots(ByVal info As Long) As Long
    Dim spectatorsJoined As Long, totalSpectatorSlots As Long
    spectatorsJoined = SHRU(info, 16) And &HFFFF&
    totalSpectatorSlots = info And &HFFFF&
    If spectatorsJoined > totalSpectatorSlots Then
        SpectatorInfo_availableSpectatorSlots = 0
        Exit Function
    End If
    SpectatorInfo_availableSpectatorSlots = totalSpectatorSlots - spectatorsJoined
End Function

Public Function updateLockedOptionsOnHost(newOpts As MultiplayOptionsLocked) As Boolean
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:264", "updateLockedOptionsOnHost", "Host only routine detected for client!": Exit Function
    If MultiplayOptionsLocked_equal(newOpts, locked) Then
        updateLockedOptionsOnHost = False
        Exit Function
    End If
    locked = newOpts
    sendHostConfig
    updateLockedOptionsOnHost = True
End Function

Public Function updateLockedOptionsFromHost(newOpts As MultiplayOptionsLocked) As Boolean
    If Not (Not NetPlay_isHost) Then ASSERT_FAILURE "!NetPlay.isHost", "src/multiint.cpp:276", "updateLockedOptionsFromHost", "Should not be called for the host": Exit Function
    If MultiplayOptionsLocked_equal(newOpts, locked) Then
        updateLockedOptionsFromHost = False
        Exit Function
    End If
    locked = newOpts
    updateLockedOptionsFromHost = True
End Function

' MultiplayOptionsLocked::operator==
Public Function MultiplayOptionsLocked_equal(a As MultiplayOptionsLocked, other As MultiplayOptionsLocked) As Boolean
    MultiplayOptionsLocked_equal = a.scavengers = other.scavengers _
        And a.alliances = other.alliances _
        And a.teams = other.teams _
        And a.power = other.power _
        And a.difficulty = other.difficulty _
        And a.ai = other.ai _
        And a.position = other.position _
        And a.bases = other.bases _
        And a.spectators = other.spectators _
        And a.name_ = other.name_ _
        And a.readybeforefull = other.readybeforefull _
        And a.cheats = other.cheats
End Function

Public Function getDifficultyListStr(ByVal idx As Long) As String
    If Not ULT(idx, 5) Then ASSERT_FAILURE "idx < difficultyList.size()", "src/multiint.cpp:287", "getDifficultyListStr", "Invalid idx: %zu", idx: Exit Function
    getDifficultyListStr = difficultyList_at(idx)
End Function

Public Function getDifficultyListCount() As Long
    getDifficultyListCount = 5
End Function

' RoomMessage::player(messageSender, messageText)
Public Function RoomMessage_player(ByVal messageSender As Long, ByVal messageText As String) As RoomMessage
    Dim message As RoomMessage
    message = RoomMessage_make(RoomMessagePlayer, messageText)
    message.sender = messageSender      ' NetPlay.playerReferences[messageSender]
    RoomMessage_player = message
End Function

Public Function RoomMessage_system(ByVal messageText As String) As RoomMessage
    RoomMessage_system = RoomMessage_make(RoomMessageSystem, messageText)
End Function

Public Function RoomMessage_notify(ByVal messageText As String) As RoomMessage
    RoomMessage_notify = RoomMessage_make(RoomMessageNotify, messageText)
End Function

' RoomMessage(RoomMessageType messageType, std::string messageText) (and its member initialisers)
Private Function RoomMessage_make(ByVal messageType As RoomMessageType, ByVal messageText As String) As RoomMessage
    RoomMessage_make.sender = -1        ' = nullptr
    RoomMessage_make.time_ = time_()    ' = std::time(nullptr)
    RoomMessage_make.typ = messageType
    RoomMessage_make.text = messageText
End Function

Public Sub displayRoomMessage(message As RoomMessage)
    Dim chatBox As WIDGET
    Set chatBox = widgGetFromID(psWScreen, MULTIOP_CHATBOX)
    If Not chatBox Is Nothing Then
        ChatBoxWidget_addMessage chatBox, message
    End If
End Sub

Public Sub displayRoomSystemMessage(ByVal text As String)
    Dim m As RoomMessage
    m = RoomMessage_system(text)
    displayRoomMessage m
End Sub

Public Sub displayRoomNotifyMessage(ByVal text As String)
    Dim m As RoomMessage
    m = RoomMessage_notify(text)
    displayRoomMessage m
End Sub

' const std::vector<WzString> getAINames(): the names in names(0 .. return value - 1)
Public Function getAINames(names() As String) As Long
    Dim i As Long, n As Long
    For i = 0 To getAIData_size() - 1
        If getAIData_js(i) <> "" Then
            ReDim Preserve names(0 To n)
            names(n) = getAIData_js(i)
            n = n + 1
        End If
    Next
    getAINames = n
End Function

Public Sub frontendCycleAIs()
    Dim currAI As String, i As Long, idx As Long
    currAI = getDefaultSkirmishAI()

    For i = 0 To getAIData_size() - 1
        If getAIData_js(i) = currAI Then
            If (i + 1) >= getAIData_size() Then idx = 0 Else idx = i + 1
            setDefaultSkirmishAI getAIData_js(idx)
            Exit For
        End If
    Next
End Sub

' ////////////////////////////////////////////////////////////////////////////
' helper func

Public Sub refreshMultiplayerOptionsTitleUIIfActive()
    Dim psCurr As WzMultiplayerOptionsTitleUI
    If TypeOf wzTitleUICurrent Is WzMultiplayerOptionsTitleUI Then
        Set psCurr = wzTitleUICurrent
        psCurr.updateGameOptions
    End If
End Sub

Public Function matchAIbyName(ByVal name_ As String) As Long
    Dim i As Long
    If Len(name_) = 0 Then
        matchAIbyName = AI_CLOSED
        Exit Function
    End If
    For i = 0 To getAIData_size() - 1
        ' strncasecmp(name, (*it).name, MAX_LEN_AI_NAME) == 0
        If StrComp(Left$(name_, MAX_LEN_AI_NAME), Left$(getAIData_name(i), MAX_LEN_AI_NAME), vbTextCompare) = 0 Then
            matchAIbyName = i
            Exit Function
        End If
    Next
    matchAIbyName = AI_NOT_FOUND
End Function

' ////////////////////////////////////////////////////////////////////////////
' Game Options Screen.

' ////////////////////////////////////////////////////////////////////////////

Public Function addInlineChooserBlueForm(ByVal psScreen As W_SCREEN, ByVal psParent As WIDGET, ByVal id As Long, ByVal txt As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As WIDGET
    Dim sFormInit As W_FORMINIT, psForm As WIDGET
    ' draw options box.
    sFormInit.formID = MULTIOP_INLINE_OVERLAY_ROOT_FRM
    sFormInit.id = id
    sFormInit.x = x And &HFFFF&
    sFormInit.y = y And &HFFFF&
    sFormInit.style = WFORM_PLAIN
    sFormInit.width = w And &HFFFF&    '190;
    sFormInit.height = h And &HFFFF&   '27;
    Set sFormInit.pDisplay = WzFunc_new(FN_intDisplayFeBox)

    ' (the std::weak_ptr psWeakParent: a strong capture, the form is not its parent's)
    Set sFormInit.calcLayout = WzFunc_new(FN_addInlineChooserBlueForm_calcLayout, x, y, psParent)
    Set psForm = widgAddForm(psScreen, sFormInit)
    If Not (Not psForm Is Nothing) Then ASSERT_FAILURE "psForm != nullptr", "src/multiint.cpp:1080", "addInlineChooserBlueForm", "widgAddForm failed": Exit Function
    Set addInlineChooserBlueForm = psForm
End Function

Public Function canChangeMapOrRandomize() As Boolean
    Dim numHumans As Long, allowed As Boolean
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:1113", "canChangeMapOrRandomize", "Host only routine detected for client!": Exit Function

    If Not ingame.localJoiningInProgress Then ' Only if game hasn't actually started yet.
        debug_ 1117, LOG_INFO, "canChangeMapOrRandomize", "Cannot randomize match options after game has started"
        canChangeMapOrRandomize = False
        Exit Function
    End If

    numHumans = NET_numHumanPlayers()
    allowed = (CSng(getLobbyChangeVoteTotal()) / CSng(numHumans)) > 0.5!

    resetLobbyChangeVoteData 'So the host can only do one change every vote session

    If numHumans = 1 Then
        canChangeMapOrRandomize = True
        Exit Function
    End If

    If Not allowed Then
        startLobbyChangeVote
        displayRoomSystemMessage "Not enough votes to randomize or change the map."
    End If

    canChangeMapOrRandomize = allowed
End Function

Public Function updatePlayerNameBox() As Boolean
    Dim pNameEditBox As WIDGET, pNameMultiBut As WIDGET, isInBlindMode As Boolean, changesAllowed As Boolean
    Dim playerNameTip As String
    Set pNameEditBox = widgGetFromID(psWScreen, MULTIOP_PNAME)
    Set pNameMultiBut = widgGetFromID(psWScreen, MULTIOP_PNAME_ICON)

    If Not (Not pNameEditBox Is Nothing) Then ASSERT_FAILURE "pNameEditBox != nullptr", "src/multiint.cpp:1145", "updatePlayerNameBox", "No player name edit box?": Exit Function
    If Not (Not pNameMultiBut Is Nothing) Then ASSERT_FAILURE "pNameMultiBut != nullptr", "src/multiint.cpp:1146", "updatePlayerNameBox", "No player name edit button?": Exit Function

    isInBlindMode = (game.blindMode <> BLIND_MODE_NONE)
    changesAllowed = ingame.hostChatPermissions(selectedPlayer) And Not isInBlindMode And Not locked.name_

    If Not isInBlindMode Then
        If changesAllowed Then
            playerNameTip = "Select Player Name"
        ElseIf Not ingame.hostChatPermissions(selectedPlayer) Then
            playerNameTip = "Player Name changes are not allowed when you are muted"
        Else
            playerNameTip = "Player Name changes are not allowed by the host"
        End If
    Else
        If game.blindMode >= BLIND_MODE_BLIND_GAME Then
            playerNameTip = "Blind Game:"
            playerNameTip = playerNameTip & vbLf
            playerNameTip = playerNameTip & "- Player names will not be visible to other players (until the game is over)"
        Else
            playerNameTip = "Blind Lobby:"
            playerNameTip = playerNameTip & vbLf
            playerNameTip = playerNameTip & "- Player names will not be visible to other players (until the game has started)"
        End If
    End If

    pNameEditBox.setTip playerNameTip

    If changesAllowed Then pNameEditBox.setState 0 Else pNameEditBox.setState WEDBS_DISABLE
    If changesAllowed Then pNameMultiBut.setState 0 Else pNameMultiBut.setState WBUT_DISABLE

    updatePlayerNameBox = True
End Function

' priorHostChatPermissions: std::array<bool, MAX_CONNECTED_PLAYERS>
Public Sub multiLobbyHandleHostOptionsChanges(priorHostChatPermissions() As Boolean)
    updatePlayerNameBox
    ' Possible future TODO: Update other UI if locked options changed?
End Sub

Public Function isHostOrAdmin() As Boolean
    isHostOrAdmin = NetPlay_isHost Or NetPlay_players_isAdmin(selectedPlayer)
End Function

' ////////////////////////////////////////////////////////////////////////////
' Colour functions

Public Function safeToUseColour(ByVal player As Long, ByVal otherPlayer As Long) As Boolean
    ' Player wants to take the colour from otherPlayer. May not take from a human otherPlayer, unless we're the host/admin.
    safeToUseColour = player = otherPlayer Or isHostOrAdmin() Or Not isHumanPlayer(otherPlayer)
End Function

Public Function playerBoxHeight(ByVal rowPosition As Long) As Long
    Dim hasPlayersTabs As Boolean, playersTop As Long, gap As Long, gapDiv As Long
    hasPlayersTabs = Not widgGetFromID(psWScreen, MULTIOP_PLAYERS_TABS) Is Nothing
    If hasPlayersTabs Then playersTop = MULTIOP_PLAYERS_TABS_H + 1 Else playersTop = 1
    gap = (MULTIOP_PLAYERSH - playersTop) - MULTIOP_TEAMSHEIGHT * numSlotsToBeDisplayed()
    gapDiv = numSlotsToBeDisplayed() - 1
    If 5 * gapDiv < gap Then gap = 5 * gapDiv
    If hasPlayersTabs Then
        gap = 0
    End If
    ' STATIC_ASSERT(MULTIOP_TEAMSHEIGHT == MULTIOP_PLAYERHEIGHT);  // Why are these different defines?
    playerBoxHeight = playersTop + (MULTIOP_TEAMSHEIGHT * gapDiv + gap) * rowPosition \ gapDiv
End Function

Public Function addMultiButWithClickHandler(ByVal parent As WIDGET, ByVal id As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal tipres As String, ByVal norm As Long, ByVal down As Long, ByVal hi As Long, ByVal clickHandler As WzFunc, Optional ByVal tc As Long = MAX_PLAYERS) As WIDGET
    Dim psButton As WIDGET
    If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/multiint.cpp:1455", "addMultiButWithClickHandler", "Null parent": Exit Function
    Set psButton = addMultiBut(parent, id, x, y, width, height, tipres, norm, down, hi, tc)
    If Not WIDGET_isKindOf(psButton, K_WzMultiButton) Then
        Exit Function
    End If
    W_BUTTON_addOnClickHandler psButton, clickHandler
    Set addMultiButWithClickHandler = psButton
End Function

Public Sub changeAI(ByVal player As Long, ByVal ai As Long)
    Dim previousPlayersShouldCheckReadyValue As Boolean
    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < MAX_CONNECTED_PLAYERS", "src/multiint.cpp:1535", "changeAI", "Invalid player: %u", player: Exit Sub

    previousPlayersShouldCheckReadyValue = multiplayPlayersShouldCheckReady()

    NetPlay_players_ai(player) = ai

    If ingame.localJoiningInProgress Then ' Only if game hasn't actually started yet.
        handlePossiblePlayersShouldCheckReadyChange previousPlayersShouldCheckReadyValue
    End If
End Sub

' std::set<uint32_t>: isValid(i) for i < MAX_PLAYERS (the set's members), their number returned
Public Function validPlayerIdxTargetsForPlayerPositionMove(ByVal player As Long, isValid() As Boolean) As Long
    Dim i As Long, n As Long
    ReDim isValid(0 To MAX_PLAYERS - 1)
    For i = 0 To game.maxPlayers - 1
        If player <> i Then
            If isHostOrAdmin() Or Not isHumanPlayer(i) Then ' host/admin can move a player to any slot, player can only move to empty slots
                If Not isSpectatorOnlySlot(i) Then ' target cannot be a spectator only slot (for player position changes)
                    isValid(i) = True
                    n = n + 1
                End If
            End If
        End If
    Next
    validPlayerIdxTargetsForPlayerPositionMove = n
End Function

Private Sub informIfAdminChangedOtherTeam(ByVal targetPlayerIdx As Long, ByVal responsibleIdx As Long)
    If Not shouldInformOfAdminAction(targetPlayerIdx, responsibleIdx) Then
        Exit Sub
    End If

    ' (sendQuickChat(INTERNAL_ADMIN_ACTION_NOTICE, ...) and the "Admin %s (%s) changed team of player ..." line
    ' with the admin's identity: the quick chat and the identities, cluster J)
End Sub

Public Sub handlePossiblePlayersShouldCheckReadyChange(ByVal previousPlayersShouldCheckReadyValue As Boolean)
    Dim currentPlayersShouldCheckReadyValue As Boolean, now_ As Currency, i As Long, shouldAccumulateNotReadyTime As Boolean
    currentPlayersShouldCheckReadyValue = multiplayPlayersShouldCheckReady()
    If previousPlayersShouldCheckReadyValue <> currentPlayersShouldCheckReadyValue Then
        If currentPlayersShouldCheckReadyValue Then
            debug_ 2431, LOG_NET, "handlePossiblePlayersShouldCheckReadyChange", "Starting to track not-ready time"

            ' Player should now check ready
            ' Update ingame.lastNotReadyTimes for all not-ready players to "now"
            QueryPerformanceCounter now_
            For i = 0 To MAX_CONNECTED_PLAYERS - 1
                If isHumanPlayer(i) Then
                    If Not NetPlay_players_ready(i) Then
                        ingame.has_lastNotReadyTimes(i) = True
                        ingame.lastNotReadyTimes(i) = now_
                    End If
                End If
            Next

            If NetPlay_isHost And bMultiPlayer And NetPlay_bComms Then
                ' (sendQuickChat(INTERNAL_LOCALIZED_LOBBY_NOTICE, ..., PlayerShouldCheckReadyNotice,
                ' war_getAutoNotReadyKickSeconds()): the quick chat, cluster J)
            End If
        Else
            debug_ 2453, LOG_NET, "handlePossiblePlayersShouldCheckReadyChange", "Stopping not-ready time tracking"
            shouldAccumulateNotReadyTime = isBlindSimpleLobby(game.blindMode)

            ' Players are no longer required to check ready
            QueryPerformanceCounter now_
            If perfFrequency = 0 Then QueryPerformanceFrequency perfFrequency
            For i = 0 To MAX_CONNECTED_PLAYERS - 1
                If isHumanPlayer(i) Then
                    If shouldAccumulateNotReadyTime Then
                        ' Update ingame.secondsNotReady for all not-ready players
                        If Not NetPlay_players_ready(i) Then
                            ' accumulate time spent not ready while they _could_ check ready
                            If ingame.has_lastNotReadyTimes(i) Then
                                ingame.secondsNotReady(i) = ingame.secondsNotReady(i) + Fix(CDbl(now_ - ingame.lastNotReadyTimes(i)) / CDbl(perfFrequency))
                            End If
                        End If
                    Else
                        ' Reset the not-ready time for all players
                        ' (time will start accumulating again when multiplayPlayersShouldCheckReady becomes true)
                        ingame.secondsNotReady(i) = 0
                    End If
                    ' Reset ingame.lastNotReadyTimes to nullopt
                    ingame.has_lastNotReadyTimes(i) = False
                End If
            Next
        End If
        wz_command_interface_output_room_status_json True
    End If
End Sub

Public Function changeTeam(ByVal player As Long, ByVal team As Long, ByVal responsibleIdx As Long) As Boolean
    Dim previousPlayersShouldCheckReadyValue As Boolean
    If team >= MAX_CONNECTED_PLAYERS Then
        changeTeam = False
        Exit Function
    End If

    If player >= MAX_CONNECTED_PLAYERS Then
        changeTeam = False
        Exit Function
    End If

    If NetPlay_players_team(player) = team Then
        NETBroadcastPlayerInfo player ' we do this regardless, in case of sync issues // FUTURE TODO: Doublecheck if this is still needed?
        changeTeam = False  ' Nothing to do.
        Exit Function
    End If

    previousPlayersShouldCheckReadyValue = multiplayPlayersShouldCheckReady()

    NetPlay_players_team(player) = team
    debug_ 2508, LOG_WZ, "changeTeam", "set %d as new team for player %d", team, player
    NETBroadcastPlayerInfo player
    informIfAdminChangedOtherTeam player, responsibleIdx

    handlePossiblePlayersShouldCheckReadyChange previousPlayersShouldCheckReadyValue

    netPlayersUpdated = True
    changeTeam = True
End Function

Public Function SendTeamRequest(ByVal player As Long, ByVal chosenTeam As Long) As Boolean
    Dim w As MessageWriter
    If NetPlay_isHost Then          ' do or request the change.
        changeTeam player, chosenTeam, realSelectedPlayer ' do the change, remember only the host can do this to avoid confusion.
    Else
        Set w = NETbeginEncode(NETnetQueue(NetPlay_hostPlayer), NET_TEAMREQUEST)
        NETuint8_t w, player
        NETuint8_t w, chosenTeam
        NETend w
    End If
    SendTeamRequest = True
End Function

Public Function recvTeamRequest(ByVal queue As Long) As Boolean
    Dim r As MessageReader, player As Long, team As Long
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:2537", "recvTeamRequest", "Host only routine detected for client!": recvTeamRequest = True: Exit Function

    Set r = NETbeginDecode(queue, NET_TEAMREQUEST)

    NETuint8_t r, player
    NETuint8_t r, team
    NETend r

    If player >= MAX_PLAYERS Or team >= MAX_PLAYERS Then
        debug_ 2548, LOG_NET, "recvTeamRequest", "NET_TEAMREQUEST invalid, player %d team, %d", player, team
        debug_ 2549, LOG_ERROR, "recvTeamRequest", "Invalid NET_TEAMREQUEST from player %d: Tried to change player %d (team %d)", _
              NETQUEUE_index(queue), player, team
        recvTeamRequest = False
        Exit Function
    End If

    If whosResponsible(player) <> NETQUEUE_index(queue) And Not NetPlay_players_isAdmin(NETQUEUE_index(queue)) Then
        HandleBadParam "NET_TEAMREQUEST given incorrect params.", player, NETQUEUE_index(queue)
        recvTeamRequest = False
        Exit Function
    End If

    If locked.teams Then
        recvTeamRequest = False
        Exit Function
    End If

    If Not alliancesSetTeamsBeforeGame(game.alliance) Then
        recvTeamRequest = False
        Exit Function
    End If

    If NetPlay_players_team(player) <> team Then
        resetReadyStatus False
    End If
    debug_ 2574, LOG_NET, "recvTeamRequest", "%s is now part of team: %d", getPlayerName(player), team
    recvTeamRequest = changeTeam(player, team, NETQUEUE_index(queue)) ' we do this regardless, in case of sync issues
End Function

Public Function recvReadyRequest(ByVal queue As Long) As Boolean
    Dim r As MessageReader, player As Long, bReady As Boolean, changedValue As Boolean
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:2609", "recvReadyRequest", "Host only routine detected for client!": recvReadyRequest = True: Exit Function

    Set r = NETbeginDecode(queue, NET_READY_REQUEST)

    bReady = False
    NETuint8_t r, player
    NETbool r, bReady
    NETend r

    If player >= MAX_CONNECTED_PLAYERS Then
        debug_ 2621, LOG_ERROR, "recvReadyRequest", "Invalid NET_READY_REQUEST from player %d: player id = %d", _
              NETQUEUE_index(queue), player
        recvReadyRequest = False
        Exit Function
    End If

    If whosResponsible(player) <> NETQUEUE_index(queue) Then
        HandleBadParam "NET_READY_REQUEST given incorrect params.", player, NETQUEUE_index(queue)
        recvReadyRequest = False
        Exit Function
    End If

    ' do not allow players to select 'ready' if we are sending a map to them!
    ' TODO: make a new icon to show this state?
    If NetPlay_players_fileSendInProgress(player) Then
        recvReadyRequest = False
        Exit Function
    End If

    ' Sanity-check: players should always have an identity to check "ready"
    If EcKey_empty(getMultiStats(player).identity) And bReady And Not NetPlay_players_isSpectator(player) Then
        ' log this!
        debug_ 2644, LOG_INFO, "recvReadyRequest", "Player has empty identity: (player: %u, name: ""%s"", IP: %s)", player, getPlayerName(player), NetPlay_players_IPtextAddress(player)

        ' kick the player
        HandleBadParam "NET_READY_REQUEST failed due to player empty identity.", player, NETQUEUE_index(queue)
        recvReadyRequest = False
        Exit Function
    End If

    If Not NetPlay_players_isSpectator(player) And Not multiplayPlayersCanCheckReady() Then
        ' silently ignore players trying to check ready when they shouldn't - refresh player info with current state
        NETBroadcastPlayerInfo player
        recvReadyRequest = False
        Exit Function
    End If

    changedValue = changeReadyStatus(player, bReady)
    If changedValue And wz_command_interface_enabled() Then
        ' (the "WZEVENT: readyStatus=%d: ..." line: the player's identity and IP address, cluster J)
        wz_command_interface_output_room_status_json
    End If
    recvReadyRequest = changedValue
End Function

Private Sub informIfAdminChangedOtherPosition(ByVal targetPlayerIdx As Long, ByVal responsibleIdx As Long)
    If Not shouldInformOfAdminAction(targetPlayerIdx, responsibleIdx) Then
        Exit Sub
    End If

    ' (sendQuickChat(INTERNAL_ADMIN_ACTION_NOTICE, ...) and the "Admin %s (%s) changed position of player ..."
    ' line with the admin's identity: the quick chat and the identities, cluster J)
End Sub

Public Function changePosition(ByVal player As Long, ByVal position As Long, ByVal responsibleIdx As Long) As Boolean
    Dim i As Long, t As Long
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:2731", "changePosition", "Host only routine detected for client!": Exit Function
    If Not ULT(player, MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/multiint.cpp:2732", "changePosition", "Invalid player idx: %u", player: Exit Function

    If NetPlay_players_position(player) = position Then
        ' nothing to do
        changePosition = False
        Exit Function
    End If

    For i = 0 To MAX_PLAYERS - 1
        If NetPlay_players_position(i) = position Then
            If Not isHumanPlayer(i) Or isPlayerHostOrAdmin(responsibleIdx) Then
                debug_ 2747, LOG_NET, "changePosition", "Swapping positions between players %d(%d) and %d(%d)", _
                      player, NetPlay_players_position(player), i, NetPlay_players_position(i)
                t = NetPlay_players_position(i)          ' std::swap(position)
                NetPlay_players_position(i) = NetPlay_players_position(player)
                NetPlay_players_position(player) = t
                t = NetPlay_players_team(i)              ' std::swap(team)
                NetPlay_players_team(i) = NetPlay_players_team(player)
                NetPlay_players_team(player) = t
                If NetPlay_players_ready(player) And isHumanPlayer(player) And ingame.JoiningInProgress(player) Then
                    changeReadyStatusImpl player, False, True
                End If
                If NetPlay_players_ready(i) And isHumanPlayer(i) And ingame.JoiningInProgress(i) Then
                    changeReadyStatusImpl i, False, True
                End If
                NETBroadcastTwoPlayerInfo player, i
                informIfAdminChangedOtherPosition player, responsibleIdx
                netPlayersUpdated = True
                changePosition = True
                Exit Function
            Else
                debug_ 2766, LOG_INFO, "changePosition", "Unable to swap positions between players %d(%d) and %d(%d) - lack of privileges", _
                      player, NetPlay_players_position(player), i, NetPlay_players_position(i)
                changePosition = False
                Exit Function
            End If
        End If
    Next
    debug_ 2772, LOG_ERROR, "changePosition", "Failed to swap positions for player %d, position %d", player, position
    If player < game.maxPlayers And position < game.maxPlayers Then
        debug_ 2775, LOG_NET, "changePosition", "corrupted positions: player (%u) new position (%u) old position (%d)", player, position, NetPlay_players_position(player)
        ' Positions were corrupted. Attempt to fix.
        NetPlay_players_position(player) = position
        If NetPlay_players_ready(player) And isHumanPlayer(player) And ingame.JoiningInProgress(player) Then
            changeReadyStatusImpl player, False, True
        End If
        NETBroadcastPlayerInfo player
        informIfAdminChangedOtherPosition player, responsibleIdx
        netPlayersUpdated = True
        changePosition = True
        Exit Function
    End If
    changePosition = False
End Function

Public Function SendColourRequest(ByVal player As Long, ByVal col As Long) As Boolean
    Dim w As MessageWriter
    If NetPlay_isHost Then          ' do or request the change
        SendColourRequest = changeColour(player, col, realSelectedPlayer)
        Exit Function
    Else
        ' clients tell the host which color they want
        Set w = NETbeginEncode(NETnetQueue(NetPlay_hostPlayer), NET_COLOURREQUEST)
        NETuint8_t w, player
        NETuint8_t w, col
        NETend w
    End If
    SendColourRequest = True
End Function

Private Sub informIfAdminChangedOtherFaction(ByVal targetPlayerIdx As Long, ByVal responsibleIdx As Long)
    If Not shouldInformOfAdminAction(targetPlayerIdx, responsibleIdx) Then
        Exit Sub
    End If

    ' (sendQuickChat(INTERNAL_ADMIN_ACTION_NOTICE, ...) and the "Admin %s (%s) changed faction of player ..."
    ' line with the admin's identity: the quick chat and the identities, cluster J)
End Sub

Public Function changeFaction(ByVal player As Long, ByVal faction_ As Long, ByVal responsibleIdx As Long) As Boolean
    If player >= MAX_PLAYERS Then
        changeFaction = False
        Exit Function
    End If

    If NetPlay_players_faction(player) = faction_ Then
        changeFaction = False  ' Nothing to do.
        Exit Function
    End If

    NetPlay_players_faction(player) = faction_
    NETBroadcastPlayerInfo player

    informIfAdminChangedOtherFaction player, responsibleIdx

    changeFaction = True
End Function

Public Function SendFactionRequest(ByVal player As Long, ByVal faction_ As Long) As Boolean
    Dim w As MessageWriter
    ' TODO: needs to be rewritten from scratch
    If Not (faction_ <= MAX_FACTION_ID) Then ASSERT_FAILURE "faction <= static_cast<UBYTE>(MAX_FACTION_ID)", "src/multiint.cpp:2906", "SendFactionRequest", "Invalid faction: %u", faction_: Exit Function
    If NetPlay_isHost Then          ' do or request the change
        changeFaction player, faction_, realSelectedPlayer
        SendFactionRequest = True
        Exit Function
    Else
        ' clients tell the host which color they want
        Set w = NETbeginEncode(NETnetQueue(NetPlay_hostPlayer), NET_FACTIONREQUEST)
        NETuint8_t w, player
        NETuint8_t w, faction_
        NETend w
    End If
    SendFactionRequest = True
End Function

Public Function SendPositionRequest(ByVal player As Long, ByVal position As Long) As Boolean
    Dim w As MessageWriter
    If NetPlay_isHost Then          ' do or request the change
        SendPositionRequest = changePosition(player, position, realSelectedPlayer)
        Exit Function
    Else
        debug_ 2931, LOG_NET, "SendPositionRequest", "Requesting the host to change our position. From %d to %d", player, position
        ' clients tell the host which position they want
        Set w = NETbeginEncode(NETnetQueue(NetPlay_hostPlayer), NET_POSITIONREQUEST)
        NETuint8_t w, player
        NETuint8_t w, position
        NETend w
    End If
    SendPositionRequest = True
End Function

Public Function recvFactionRequest(ByVal queue As Long) As Boolean
    Dim r As MessageReader, player As Long, faction_ As Long, newFactionId As Long
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:2943", "recvFactionRequest", "Host only routine detected for client!": recvFactionRequest = True: Exit Function

    Set r = NETbeginDecode(queue, NET_FACTIONREQUEST)

    NETuint8_t r, player
    NETuint8_t r, faction_
    NETend r

    If player >= MAX_PLAYERS Then
        debug_ 2954, LOG_ERROR, "recvFactionRequest", "Invalid NET_FACTIONREQUEST from player %d: Tried to change player %d to faction %d", _
              NETQUEUE_index(queue), player, faction_
        recvFactionRequest = False
        Exit Function
    End If

    If whosResponsible(player) <> NETQUEUE_index(queue) And Not NetPlay_players_isAdmin(NETQUEUE_index(queue)) Then
        HandleBadParam "NET_FACTIONREQUEST given incorrect params.", player, NETQUEUE_index(queue)
        recvFactionRequest = False
        Exit Function
    End If

    If Not uintToFactionID(faction_, newFactionId) Then
        HandleBadParam "NET_FACTIONREQUEST given incorrect params.", player, NETQUEUE_index(queue)
        recvFactionRequest = False
        Exit Function
    End If

    recvFactionRequest = changeFaction(player, newFactionId, NETQUEUE_index(queue))
End Function

Public Function recvColourRequest(ByVal queue As Long) As Boolean
    Dim r As MessageReader, player As Long, col As Long
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:2977", "recvColourRequest", "Host only routine detected for client!": recvColourRequest = True: Exit Function

    Set r = NETbeginDecode(queue, NET_COLOURREQUEST)

    NETuint8_t r, player
    NETuint8_t r, col
    NETend r

    If player >= MAX_PLAYERS Then
        debug_ 2988, LOG_ERROR, "recvColourRequest", "Invalid NET_COLOURREQUEST from player %d: Tried to change player %d to colour %d", _
              NETQUEUE_index(queue), player, col
        recvColourRequest = False
        Exit Function
    End If

    If whosResponsible(player) <> NETQUEUE_index(queue) And Not NetPlay_players_isAdmin(NETQUEUE_index(queue)) Then
        HandleBadParam "NET_COLOURREQUEST given incorrect params.", player, NETQUEUE_index(queue)
        recvColourRequest = False
        Exit Function
    End If

    recvColourRequest = changeColour(player, col, NETQUEUE_index(queue))
End Function

Public Function recvPositionRequest(ByVal queue As Long) As Boolean
    Dim r As MessageReader, player As Long, position As Long
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:3004", "recvPositionRequest", "Host only routine detected for client!": recvPositionRequest = True: Exit Function

    Set r = NETbeginDecode(queue, NET_POSITIONREQUEST)

    NETuint8_t r, player
    NETuint8_t r, position
    NETend r
    debug_ 3012, LOG_NET, "recvPositionRequest", "Host received position request from player %d to %d", player, position

    If player >= MAX_PLAYERS Or position >= MAX_PLAYERS Then
        debug_ 3016, LOG_ERROR, "recvPositionRequest", "Invalid NET_POSITIONREQUEST from player %d: Tried to change player %d to %d", _
              NETQUEUE_index(queue), player, position
        recvPositionRequest = False
        Exit Function
    End If

    If whosResponsible(player) <> NETQUEUE_index(queue) And Not NetPlay_players_isAdmin(NETQUEUE_index(queue)) Then
        HandleBadParam "NET_POSITIONREQUEST given incorrect params.", player, NETQUEUE_index(queue)
        recvPositionRequest = False
        Exit Function
    End If

    If locked.position Then
        recvPositionRequest = False
        Exit Function
    End If

    If Not changePosition(player, position, NETQUEUE_index(queue)) Then
        recvPositionRequest = False
        Exit Function
    End If

    wz_command_interface_output_room_status_json True
    recvPositionRequest = True
End Function

Public Function SendPlayerSlotTypeRequest(ByVal player As Long, ByVal isSpectator As Boolean) As Boolean
    Dim originalPlayerSlotType As String, desiredPlayerSlotType As String, w As MessageWriter, q As Long
    If NetPlay_players_isSpectator(player) Then originalPlayerSlotType = "Spectator" Else originalPlayerSlotType = "Player"
    If isSpectator Then desiredPlayerSlotType = "Spectator" Else desiredPlayerSlotType = "Player"

    ' If I'm a spectator, and I ask to be a player, record locally that I gave permission for such a change
    If Not NetPlay_isHost And NetPlay_players_isSpectator(player) And Not isSpectator Then
        bRequestedSelfMoveToPlayers = True
    End If
    If NetPlay_isHost And NetPlay_players_isSpectator(player) And Not isSpectator Then
        ' If the host asks the spectator if they want to be a player, record this so we can appropriately flag that later move as "host initiated"
        bHostRequestedMoveToPlayers(player) = True
    End If

    debug_ 3057, LOG_NET, "SendPlayerSlotTypeRequest", "Requesting the host to change our slot type. From %s to %s", originalPlayerSlotType, desiredPlayerSlotType
    ' clients tell the host which player slot type they want (but the host may not allow)
    If Not NetPlay_isHost Then q = NetPlay_hostPlayer Else q = player
    Set w = NETbeginEncode(NETnetQueue(q), NET_PLAYER_SLOTTYPE_REQUEST)
    NETuint32_t w, player
    NETbool w, isSpectator
    NETend w
    SendPlayerSlotTypeRequest = True
End Function

Public Function recvPlayerSlotTypeRequestAndPop(ByVal titleUI As WzMultiplayerOptionsTitleUI, ByVal queue As Long) As Boolean
    Dim playerIndex As Long, desiredIsSpectator As Boolean, r As MessageReader, playerName As String, text As String
    Dim pPlayerName As String, msg As String, result As SpectatorToPlayerMoveResult, notificationTag As String
    Dim has_swapChooserUp As Boolean, swapChooserUp As Long
    ' A player is requesting a slot type change
    ' ex. player -> spectator

    desiredIsSpectator = False
    Set r = NETbeginDecode(queue, NET_PLAYER_SLOTTYPE_REQUEST)
    NETuint32_t r, playerIndex
    NETbool r, desiredIsSpectator
    NETend r

    NETpop queue ' this function *must* handle popping the message itself since, if a player index switch occurs, the queue may be invalidated

    ' Basic parameter validation

    If Not ingame.localJoiningInProgress Then ' Only if game hasn't actually started yet.
        debug_ 3084, LOG_WARNING, "recvPlayerSlotTypeRequestAndPop", "Player indexes cannot be swapped after game has started"
        recvPlayerSlotTypeRequestAndPop = False
        Exit Function
    End If

    If playerIndex >= MAX_CONNECTED_PLAYERS Then
        debug_ 3090, LOG_ERROR, "recvPlayerSlotTypeRequestAndPop", "Invalid NET_PLAYER_SLOTTYPE_REQUEST from player %u: Tried to change player %u", _
              NETQUEUE_index(queue), playerIndex
        recvPlayerSlotTypeRequestAndPop = False
        Exit Function
    End If

    If whosResponsible(playerIndex) <> NETQUEUE_index(queue) And Not (NETQUEUE_index(queue) = NetPlay_hostPlayer And whosResponsible(playerIndex) = selectedPlayer And selectedPlayer <> NetPlay_hostPlayer) Then
        HandleBadParam "NET_PLAYER_SLOTTYPE_REQUEST given incorrect params.", playerIndex, NETQUEUE_index(queue)
        recvPlayerSlotTypeRequestAndPop = False
        Exit Function
    End If

    If desiredIsSpectator = NetPlay_players_isSpectator(playerIndex) Then
        ' no-op change - this may be a response to a request
        playerName = getPlayerName(playerIndex)
        If NetPlay_isHost Then
            ' client sending response
            If desiredIsSpectator Then
                text = va("Spectator %s wants to remain a Spectator", playerName)
            Else
                text = va("Player %s wants to remain a Player", playerName)
            End If
        Else
            ' host sending a denial
            If desiredIsSpectator Then
                text = "Host has declined to switch you to a Player"
            Else
                text = "Unable to switch to Spectator"
            End If
        End If
        displayRoomSystemMessage text
        recvPlayerSlotTypeRequestAndPop = False
        Exit Function
    End If

    If Not NetPlay_isHost Then
        ' Client-only message handling
        ' The host can ask a client (us) if we want to switch player slot type
        If NETQUEUE_index(queue) <> NetPlay_hostPlayer Then
            ' Only the host can ask!
            debug_ 3141, LOG_ERROR, "recvPlayerSlotTypeRequestAndPop", "Invalid NET_PLAYER_SLOTTYPE_REQUEST from player %u", NETQUEUE_index(queue)
            recvPlayerSlotTypeRequestAndPop = False
            Exit Function
        End If

        If Not hasNotificationsWithTag(SLOTTYPE_REQUEST_TAG) Then
            ' (the WZ_Notification "Do you want to spectate?" / "Do you want to play?" with its action
            ' SendPlayerSlotTypeRequest(selectedPlayer, desiredIsSpectator) and its dismissal / ignore answers:
            ' notifications.cpp, cluster J)
        End If

        recvPlayerSlotTypeRequestAndPop = True
        Exit Function
    End If

    ' Host-only message handling

    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:3198", "recvPlayerSlotTypeRequestAndPop", "Host only routine detected for client!": recvPlayerSlotTypeRequestAndPop = False: Exit Function

    If locked.spectators Then
        recvPlayerSlotTypeRequestAndPop = False
        Exit Function
    End If

    pPlayerName = getPlayerName(playerIndex, True)
    If Len(pPlayerName) <> 0 Then playerName = pPlayerName Else playerName = "[p" & CStr(playerIndex) & "]"

    If desiredIsSpectator Then
        If Not NETmovePlayerToSpectatorOnlySlot(playerIndex, False) Then
            ' Notify the player that the answer is "no" / the move failed by sending a no-op PlayerSlotTypeRequest
            SendPlayerSlotTypeRequest playerIndex, NetPlay_players_isSpectator(playerIndex)
            recvPlayerSlotTypeRequestAndPop = False
            Exit Function
        End If

        msg = va("Player %s has moved to Spectators", playerName)
        sendRoomSystemMessage msg
        resetReadyStatus True       'reset and send notification to all clients
    Else
        ' Spectator wants to switch to a player slot

        ' Try a move to any open player slot
        result = NETmoveSpectatorToPlayerSlot(playerIndex, -1, bHostRequestedMoveToPlayers(playerIndex))
        bHostRequestedMoveToPlayers(playerIndex) = False
        Select Case result
        Case SpectatorToPlayerMoveResult_SUCCESS
            ' Was able to move spectator to an open player slot automatically
            msg = va("Spectator %s has moved to Players", playerName)
            sendRoomSystemMessage msg
            resetReadyStatus True       'reset and send notification to all clients
        Case SpectatorToPlayerMoveResult_NEEDS_SLOT_SELECTION
            If headlessGameMode() Then
                ' displaying a UI won't work
                ' so instead, send a room message about the failure
                msg = va("Unable to move %s to Players - no available slot", playerName)
                sendRoomSystemMessage msg
                recvPlayerSlotTypeRequestAndPop = True ' and immediately return
                Exit Function
            End If

            ' Display a notification that a spectator would like to switch to a player
            notificationTag = SLOTTYPE_REQUEST_TAG
            notificationTag = notificationTag & "::" & CStr(playerIndex)
            If hasNotificationsWithTag(notificationTag) Then
                ' no nagging permitted - there's already a notification queued for this player index
            Else
                has_swapChooserUp = titleUI.playerSlotSwapChooserOpenForPlayer(swapChooserUp)
                If has_swapChooserUp And swapChooserUp = playerIndex Then
                    ' no nagging permitted - the host already has the player slot swap chooser up for this player!
                Else
                    ' (the WZ_Notification "Spectator would like to become a Player", whose action schedules
                    ' titleUI->openPlayerSlotSwapChooser(playerIndex) and whose dismissal answers "no" with a no-op
                    ' SendPlayerSlotTypeRequest: notifications.cpp, cluster J)
                    ' "handled" - now up to the host to decide what happens, so return true below
                End If
            End If
        Case SpectatorToPlayerMoveResult_FAILED
            ' Notify the player that the answer is "no" by sending a no-op PlayerSlotTypeRequest
            SendPlayerSlotTypeRequest playerIndex, NetPlay_players_isSpectator(playerIndex)
            recvPlayerSlotTypeRequestAndPop = False
            Exit Function
        End Select
    End If

    recvPlayerSlotTypeRequestAndPop = True
End Function

' (VB) client-only (the host never receives it): its checks; the swap itself (the player references, NET_InitPlayer,
' the stats and mute swaps, the simulated re-join) is the joining's, cluster J
Public Function recvSwapPlayerIndexes(ByVal queue As Long, ByVal parent As WzTitleUI) As SwapPlayerIndexesResult
    Dim playerIndexA As Long, playerIndexB As Long, r As MessageReader
    ' A special message that can be triggered by the host to queue-up a player index swap (for example, converting a player to a spectator-only slot - really should only be used for converting between player and spectator slots in the lobby)
    ' For the player being moved, this simulates a lot of the effects of a join, followed by a (quiet / virtual) "leave" of the old slot
    ' For other players, this provides a message and a (quiet / virtual) "leave" of the old slot
    ' Everyone then waits for the host to provide updated NET_PLAYER_INFO

    If Not (NETQUEUE_index(queue) = NetPlay_hostPlayer) Then ASSERT_FAILURE "queue.index == NetPlay.hostPlayer", "src/multiint.cpp:3321", "recvSwapPlayerIndexes", "Message received from non-host user": recvSwapPlayerIndexes = SwapPlayerIndexesResult_FAILURE: Exit Function

    Set r = NETbeginDecode(queue, NET_PLAYER_SWAP_INDEX)
    NETuint32_t r, playerIndexA
    NETuint32_t r, playerIndexB
    NETend r

    If Not ingame.localJoiningInProgress Then ' Only if game hasn't actually started yet.
        debug_ 3333, LOG_ERROR, "recvSwapPlayerIndexes", "Player indexes cannot be swapped after game has started"
        recvSwapPlayerIndexes = SwapPlayerIndexesResult_FAILURE
        Exit Function
    End If
    If NetPlay_isHost Then
        debug_ 3338, LOG_ERROR, "recvSwapPlayerIndexes", "Should never be called for the host"
        recvSwapPlayerIndexes = SwapPlayerIndexesResult_FAILURE
        Exit Function
    End If

    If playerIndexA >= MAX_CONNECTED_PLAYERS Then
        debug_ 3344, LOG_ERROR, "recvSwapPlayerIndexes", "Bad player number (%u) received from host!", playerIndexA
        recvSwapPlayerIndexes = SwapPlayerIndexesResult_FAILURE
        Exit Function
    End If

    If playerIndexB >= MAX_CONNECTED_PLAYERS Then
        debug_ 3350, LOG_ERROR, "recvSwapPlayerIndexes", "Bad player number (%u) received from host!", playerIndexB
        recvSwapPlayerIndexes = SwapPlayerIndexesResult_FAILURE
        Exit Function
    End If

    ' Make sure neither is the host index, as this is *not* supported for the host
    If playerIndexA = NetPlay_hostPlayer Or playerIndexB = NetPlay_hostPlayer Then
        debug_ 3357, LOG_ERROR, "recvSwapPlayerIndexes", "Cannot swap host slot! (players: %u, %u)!", playerIndexA, playerIndexB
        recvSwapPlayerIndexes = SwapPlayerIndexesResult_FAILURE
        Exit Function
    End If

    ' (the swap: cluster J)
    recvSwapPlayerIndexes = SwapPlayerIndexesResult_FAILURE
End Function

' host kicks a player from a game.
Public Sub kickPlayer(ByVal player_id As Long, ByVal reason As String, ByVal typ As LOBBY_ERROR_TYPES, Optional ByVal banPlayer As Boolean = False)
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    Dim w As MessageWriter, t As Long
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:3848", "kickPlayer", "Host only routine detected for client!": Exit Sub

    debug_ 3850, LOG_INFO, "kickPlayer", "Kicking player %u (%s). Reason: %s", player_id, getPlayerName(player_id), reason

    ' send a kick msg
    Set w = NETbeginEncode(NETbroadcastQueue(), NET_KICK)
    NETuint32_t w, player_id
    NETstring w, reason, MAX_KICK_REASON
    t = typ
    NETenum w, t
    NETend w
    NETflush
    wzDelay 300

    ActivityManager_hostKickPlayer player_id, typ, reason

    If banPlayer And NetPlay_players_allocated(player_id) Then
        addIPToBanList NetPlay_players_IPtextAddress(player_id), NetPlay_players_name(player_id)
    End If

    NETplayerKicked player_id
End Sub

Public Function buildMessage(ByVal sender As Long, ByVal text As String) As RoomMessage
    Select Case sender
    Case SYSTEM_MESSAGE
        buildMessage = RoomMessage_system(text)
    Case NOTIFY_MESSAGE
        buildMessage = RoomMessage_notify(text)
    Case Else
        If sender >= 0 And sender < MAX_CONNECTED_PLAYERS Then
            buildMessage = RoomMessage_player(sender, text)
            Exit Function
        End If

        debug_ 3954, LOG_ERROR, "buildMessage", "Invalid message sender %d.", sender
        buildMessage = RoomMessage_system(text)
    End Select
End Function

'--------------------------------------------------------------------------
' class ChatBoxButton : public W_BUTTON

Public Sub ChatBoxButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_ChatBoxButton ' (VB) virtual calls from here on reach ChatBoxButton's overrides
    Set ChatBoxButton_(self).text = New WzText
End Sub

Public Function ChatBoxButton_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_ChatBoxButton)
    ChatBoxButton_construct w
    Set ChatBoxButton_new = w
End Function

Public Sub ChatBoxButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, haveText As Boolean, isDown As Boolean, isDisabled As Boolean
    Dim isHighlight As Boolean, border_color As Long, fill_color As Long, fw As Long, fx As Long, fy As Long, textColor As Long
    Dim text As WzText
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    haveText = Len(W_BUTTON_(self).pText) <> 0

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    ' Display the button.
    If Not isDisabled Then border_color = pal_RGBA(255, 255, 255, 120) Else border_color = psPalette(WZCOL_FORM_DARK)
    If isDown Or isDisabled Then fill_color = pal_RGBA(10, 0, 70, 250) Else fill_color = pal_RGBA(25, 0, 110, 220)
    iV_ShadowBox x0, y0, x1, y1, 0, border_color, border_color, fill_color
    If isHighlight Then
        iV_Box x0 + 2, y0 + 2, x1 - 2, y1 - 2, psPalette(WZCOL_FORM_HILITE)
    End If

    If haveText Then
        Set text = ChatBoxButton_(self).text
        text.setText W_BUTTON_(self).pText, W_BUTTON_(self).FontID
        fw = text.width()
        fx = x0 + (self.width() - fw) \ 2
        fy = y0 + (self.height() - text.lineSize()) \ 2 - text.aboveBase()
        textColor = psPalette(WZCOL_FORM_TEXT)
        If isDisabled Then
            textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
        End If
        text.render fx, fy, textColor
    End If
End Sub

' (VB) a PIELIGHT's alpha (its byte.a: the top byte)
Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

'--------------------------------------------------------------------------
' class ChatBoxWidget : public IntFormAnimated (its std::weak_ptr<ChatBoxWidget>: the widgweak_vb serial in its
' data's weakSelf)

' ChatBoxWidget(): IntFormAnimated(true) {} (and the members' initialisers)
Public Sub ChatBoxWidget_construct(ByVal self As WIDGET)
    IntFormAnimated_construct self, True
    self.kind = K_ChatBoxWidget ' (VB) virtual calls from here on reach ChatBoxWidget's overrides
    With ChatBoxWidget_(self)
        .currentMode = ChatBoxSendMode_ENABLED
        .hasFocus = False
        .weakSelf = WeakWidget_register(self)
    End With
End Sub

' static std::shared_ptr<ChatBoxWidget> make()
Public Function ChatBoxWidget_make() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_ChatBoxWidget)
    ChatBoxWidget_construct widget
    vt_ChatBoxWidget_initialize widget
    Set ChatBoxWidget_make = widget
End Function

Public Sub ChatBoxWidget_initialize(ByVal self As WIDGET)
    Dim d As ChatBoxWidget, weakSelf As Long, minButtonWidthForText As Long, sEdInit As W_EDBINIT, pad As Padding
    Dim newRootFrm As WIDGET, psWeakChatBoxWidget As Long
    Set d = ChatBoxWidget_(self)
    self.id = MULTIOP_CHATBOX
    weakSelf = d.weakSelf

    Set d.messages = ScrollableListWidget_make()
    self.attach d.messages
    ScrollableListWidget_setSnapOffset d.messages, True
    ScrollableListWidget_setStickToBottom d.messages, True
    pad.top = 3: pad.right_ = 4: pad.bottom = 3: pad.left_ = 4
    ScrollableListWidget_setPadding d.messages, pad
    ScrollableListWidget_setItemSpacing d.messages, 1

    ' ([&]: this, the chat box: its serial, the listener being removed by the destructor)
    Set d.handleConsoleMessage = WzFunc_new(FN_chatBox_consoleListener, weakSelf)

    Set d.quickChatButton = ChatBoxButton_new()
    self.attach d.quickChatButton
    ChatBoxWidget_setQuickChatButtonDisplay self
    W_BUTTON_addOnClickHandler d.quickChatButton, WzFunc_new(FN_chatBox_quickChat_onClick, weakSelf)

    If NetPlay_bComms Then
        Set d.optionsButton = ChatBoxButton_new()
        self.attach d.optionsButton
        d.optionsButton.setString ChrW$(&H2699) ' "(gear)"
        d.optionsButton.setTip "Chat Options"
        W_BUTTON_(d.optionsButton).FontID = font_regular
        minButtonWidthForText = iV_GetTextWidth(W_BUTTON_(d.optionsButton).pText, W_BUTTON_(d.optionsButton).FontID)
        d.optionsButton.setGeometry 0, 0, minButtonWidthForText + 10, iV_GetTextLineSize(W_BUTTON_(d.optionsButton).FontID)
        W_BUTTON_addOnClickHandler d.optionsButton, WzFunc_new(FN_chatBox_options_onClick, weakSelf)
    End If

    sEdInit.formID = MULTIOP_CHATBOX
    sEdInit.id = MULTIOP_CHATEDIT
    sEdInit.pUserData = Empty
    Set sEdInit.pBoxDisplay = WzFunc_new(FN_displayChatEdit)
    Set d.editBox = W_EDITBOX_new_init(sEdInit)
    self.attach d.editBox
    ChatBoxWidget_setEditBoxDisplay self
    W_EDITBOX_setOnReturnHandler d.editBox, WzFunc_new(FN_chatBox_onEditReturn, weakSelf)
    W_EDITBOX_setOnTabHandler d.editBox, WzFunc_new(FN_chatBox_onTab, weakSelf)

    consoleAddMessageListener d.handleConsoleMessage

    ' Initialize the overlay screen
    Set d.quickChatOverlayScreen = W_SCREEN_make()
    Set newRootFrm = W_FULLSCREENOVERLAY_CLICKFORM_make()
    psWeakChatBoxWidget = weakSelf
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_chatBox_overlay_onClicked, psWeakChatBoxWidget)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onCancelPressed = W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc
    d.quickChatOverlayScreen.psForm.attach newRootFrm

    ' Add the Quick Chat form to the overlay screen form
    Set d.quickChatForm = createQuickChatForm(WzQuickChatContext_Lobby, WzFunc_new(FN_chatBox_quickChatSent, psWeakChatBoxWidget))
    newRootFrm.attach d.quickChatForm
    d.quickChatForm.setCalcLayout WzFunc_new(FN_chatBox_quickChatForm_calcLayout, psWeakChatBoxWidget)

    ChatBoxWidget_handleUpdateInHostedLobby self
End Sub

Public Sub ChatBoxWidget_handleUpdateInHostedLobby(ByVal self As WIDGET)
    Dim d As ChatBoxWidget
    Set d = ChatBoxWidget_(self)
    If Not d.optionsButton Is Nothing Then
        If bInActualHostedLobby Then
            d.optionsButton.setState 0
        Else
            d.optionsButton.setState WBUT_DISABLE
        End If
    End If
End Sub

Private Sub ChatBoxWidget_setQuickChatButtonDisplay(ByVal self As WIDGET)
    Dim d As ChatBoxWidget, compactButtonForm As Boolean, str_ As String, minButtonWidthForText As Long
    Set d = ChatBoxWidget_(self)
    If Not (Not d.quickChatButton Is Nothing) Then ASSERT_FAILURE "quickChatButton != nullptr", "src/multiint.cpp:4106", "setQuickChatButtonDisplay", "No quick chat button?": Exit Sub
    compactButtonForm = (d.currentMode = ChatBoxSendMode_ENABLED)
    If compactButtonForm Then
        d.quickChatButton.setString "+"
    Else
        str_ = "+ " & "Quick Chat"
        d.quickChatButton.setString str_
    End If
    d.quickChatButton.setTip "Quick Chat"
    W_BUTTON_(d.quickChatButton).FontID = font_regular_bold
    minButtonWidthForText = iV_GetTextWidth(W_BUTTON_(d.quickChatButton).pText, W_BUTTON_(d.quickChatButton).FontID)
    d.quickChatButton.setGeometry d.quickChatButton.x(), d.quickChatButton.y(), minButtonWidthForText + 20, iV_GetTextLineSize(W_BUTTON_(d.quickChatButton).FontID)
End Sub

Private Sub ChatBoxWidget_setEditBoxDisplay(ByVal self As WIDGET)
    Dim d As ChatBoxWidget
    Set d = ChatBoxWidget_(self)
    Select Case d.currentMode
    Case ChatBoxSendMode_DISABLED
        W_EDITBOX_setPlaceholder d.editBox, "Press the Tab key to open Quick Chat."
        W_EDITBOX_setPlaceholderTextColor d.editBox, True, psPalette(WZCOL_TEXT_MEDIUM)
        d.editBox.setTip "The host has disabled free chat. Please use Quick Chat."
        d.editBox.setState WEDBS_DISABLE
    Case ChatBoxSendMode_HOSTMSG_ONLY
        W_EDITBOX_setPlaceholder d.editBox, "Press the Tab key to open Quick Chat."
        W_EDITBOX_setPlaceholderTextColor d.editBox, True, psPalette(WZCOL_TEXT_MEDIUM)
        d.editBox.setTip "The host has disabled free chat. Please use Quick Chat or /hostmsg commands."
        d.editBox.setState 0
    Case ChatBoxSendMode_ENABLED
        W_EDITBOX_setPlaceholder d.editBox, ""
        W_EDITBOX_setPlaceholderTextColor d.editBox, False    ' nullopt
        d.editBox.setTip ""
        d.editBox.setState 0
    End Select
End Sub

Public Sub ChatBoxWidget_setSendMode(ByVal self As WIDGET, ByVal mode As ChatBoxSendMode)
    Dim d As ChatBoxWidget
    Set d = ChatBoxWidget_(self)
    If mode = d.currentMode Then
        Exit Sub
    End If
    d.currentMode = mode
    ' Update UI
    ChatBoxWidget_setQuickChatButtonDisplay self
    ChatBoxWidget_setEditBoxDisplay self
    self.geometryChanged
    ' message the user
    Select Case mode
    Case ChatBoxSendMode_DISABLED
        displayRoomSystemMessage "The host has disabled free chat. Please use Quick Chat."
    Case ChatBoxSendMode_HOSTMSG_ONLY
        displayRoomSystemMessage "The host has disabled free chat. Please use Quick Chat or /hostmsg commands."
    Case ChatBoxSendMode_ENABLED
        displayRoomSystemMessage "The host has enabled free chat for you."
    End Select
End Sub

Private Function ChatBoxWidget_calculateQuickChatScreenPosY(ByVal self As WIDGET) As Long
    Dim idealQuickChatHeight As Long, screenY0 As Long
    idealQuickChatHeight = ChatBoxWidget_(self).quickChatForm.idealHeight()
    ' Try to position it above the chat widget
    screenY0 = self.screenPosY() - idealQuickChatHeight
    ChatBoxWidget_calculateQuickChatScreenPosY = MAX_(screenY0, 0)
End Function

Public Sub ChatBoxWidget_takeFocus(ByVal self As WIDGET)
    Dim d As ChatBoxWidget, context As W_CONTEXT, lockedScreen As W_SCREEN
    Set d = ChatBoxWidget_(self)
    If d.currentMode <> ChatBoxSendMode_DISABLED Then
        ' give the edit box the focus
        ' (W_CONTEXT::ZeroContext())
        context.mx = mouseX()
        context.my = mouseY()
        W_EDITBOX_simulateClick d.editBox, context, True
    Else
        ' when the edit box is disabled, keyboard events (tab, enter) should open the quick chat
        ' to handle this, tell the form that the ChatBoxWidget has focus
        Set lockedScreen = self.screenPointer()
        If Not lockedScreen Is Nothing Then
            lockedScreen.setFocus self
            d.hasFocus = True
        Else
            ' If the ChatBoxWidget isn't currently attached to a screen when this is triggered, focus issues may occur
            ASSERT_FAILURE "false", "src/multiint.cpp:4204", "takeFocus", "ChatBoxWidget is not attached to any screen?"
        End If
    End If
End Sub

Public Sub ChatBoxWidget_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As ChatBoxWidget
    Set d = ChatBoxWidget_(self)
    If Not d.hasFocus Then Exit Sub

    If keyPressed(KEY_TAB) Then
        widgScheduleTask WzFunc_new(FN_chatBox_openQuickChat_task, d.weakSelf)
        inputLoseFocus
    End If
    If keyPressed(KEY_ESC) Then
        ' "eat" the ESC keypress, just as if the editbox had focus
        inputLoseFocus
    End If
End Sub

Public Sub ChatBoxWidget_focusLost(ByVal self As WIDGET)
    ChatBoxWidget_(self).hasFocus = False
End Sub

Private Sub ChatBoxWidget_openQuickChatOverlay(ByVal self As WIDGET)
    Dim d As ChatBoxWidget, context As W_CONTEXT
    Set d = ChatBoxWidget_(self)
    d.quickChatOverlayScreen.screenSizeDidChange screenWidth, screenHeight, screenWidth, screenHeight ' trigger a screenSizeDidChange so it can relayout (if screen size changed since last time it was registered as an overlay screen)
    widgRegisterOverlayScreen d.quickChatOverlayScreen, 1
    W_EDITBOX_stopEditing d.editBox

    ' Set the focus to the quickChatForm by simulating a click on it
    ' (W_CONTEXT::ZeroContext())
    context.mx = d.quickChatForm.screenPosX()
    context.my = d.quickChatForm.screenPosY()
    d.quickChatForm.clicked context, WKEY_PRIMARY
    d.quickChatForm.released context, WKEY_PRIMARY
End Sub

Private Sub ChatBoxWidget_closeQuickChatOverlay(ByVal self As WIDGET)
    Dim d As ChatBoxWidget, sContext As W_CONTEXT
    Set d = ChatBoxWidget_(self)
    widgRemoveOverlayScreen d.quickChatOverlayScreen
    ' Auto-click edit box
    ' (W_CONTEXT::ZeroContext())
    d.editBox.clicked sContext
End Sub

' ChatBoxWidget::~ChatBoxWidget()
Public Sub ChatBoxWidget_destruct(ByVal self As WIDGET)
    Dim d As ChatBoxWidget
    Set d = ChatBoxWidget_(self)
    consoleRemoveMessageListener d.handleConsoleMessage
    widgRemoveOverlayScreen d.quickChatOverlayScreen
    If Not d.optionsOverlayScreen Is Nothing Then
        widgRemoveOverlayScreen d.optionsOverlayScreen
    End If
    WeakWidget_unregister d.weakSelf
End Sub

Private Sub ChatBoxWidget_displayOptionsOverlay(ByVal self As WIDGET, ByVal psParent As WIDGET)
    Dim d As ChatBoxWidget, lockedScreen As W_SCREEN, newRootFrm As WIDGET, psWeakOptionsOverlayScreen As Long
    Dim psWeakChatBoxWidget As Long, chatOptionsForm As WIDGET
    Set d = ChatBoxWidget_(self)
    Set lockedScreen = self.screenPointer()
    If Not (Not lockedScreen Is Nothing) Then ASSERT_FAILURE "lockedScreen != nullptr", "src/multiint.cpp:4270", "displayOptionsOverlay", "The ChatBoxWidget does not have an associated screen pointer?"

    ' Initialize the options overlay screen
    Set d.optionsOverlayScreen = W_SCREEN_make()
    Set newRootFrm = W_FULLSCREENOVERLAY_CLICKFORM_make()
    psWeakOptionsOverlayScreen = ObjPtr(d.optionsOverlayScreen)
    psWeakChatBoxWidget = d.weakSelf
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_chatBox_optionsOverlay_onClicked, psWeakOptionsOverlayScreen, psWeakChatBoxWidget)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onCancelPressed = W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc
    d.optionsOverlayScreen.psForm.attach newRootFrm

    ' Add the Chat Options form to the overlay screen form
    Set chatOptionsForm = createChatOptionsForm(NetPlay_isHost, WzFunc_new(FN_chatBox_chatOptionsChanged))
    newRootFrm.attach chatOptionsForm
    chatOptionsForm.setCalcLayout WzFunc_new(FN_chatBox_chatOptionsForm_calcLayout, psWeakChatBoxWidget)

    widgRegisterOverlayScreenOnTopOfScreen d.optionsOverlayScreen, lockedScreen
End Sub

Public Sub ChatBoxWidget_addMessage(ByVal self As WIDGET, message As RoomMessage)
    ReDim Preserve persistentMessageLocalStorage(0 To persistentMessageLocalStorage_size)
    persistentMessageLocalStorage(persistentMessageLocalStorage_size) = message
    persistentMessageLocalStorage_size = persistentMessageLocalStorage_size + 1
    ChatBoxWidget_displayMessage self, message
End Sub

Public Sub ChatBoxWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As ChatBoxWidget, messagesHeight As Long, messageEditAreaY0 As Long, buttonsWidth As Long, chatEditBoxWidth As Long
    Set d = ChatBoxWidget_(self)
    messagesHeight = self.height() - MULTIOP_CHATEDITAREA_HEIGHT - 1
    d.messages.setGeometry 1, 1, self.width() - 2, messagesHeight

    messageEditAreaY0 = d.messages.y() + messagesHeight + MULTIOP_CHATEDITAREA_INSET

    d.quickChatButton.setGeometry MULTIOP_CHATEDITAREA_INSET, messageEditAreaY0, d.quickChatButton.width(), MULTIOP_CHATEDITH
    If Not d.optionsButton Is Nothing Then
        d.optionsButton.setGeometry self.width() - MULTIOP_CHATEDITAREA_INSET - d.optionsButton.width(), messageEditAreaY0, d.optionsButton.width(), MULTIOP_CHATEDITH
    End If
    buttonsWidth = d.quickChatButton.width()
    If Not d.optionsButton Is Nothing Then buttonsWidth = buttonsWidth + d.optionsButton.width()

    chatEditBoxWidth = self.width() - (MULTIOP_CHATEDITAREA_INSET * 2) - buttonsWidth - (MULTIOP_CHATEDITAREA_INTERNAL_PADDING * 2)
    d.editBox.setGeometry d.quickChatButton.x() + d.quickChatButton.width() + MULTIOP_CHATEDITAREA_INTERNAL_PADDING, messageEditAreaY0, chatEditBoxWidth, MULTIOP_CHATEDITH
End Sub

Public Sub ChatBoxWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As ChatBoxWidget, x0 As Long, y0 As Long, messageEditAreaY0 As Long, messageEditAreaX0 As Long
    Set d = ChatBoxWidget_(self)
    IntFormAnimated_display self, xOffset, yOffset

    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    ' draws the line at the bottom of the multiplayer join dialog separating the chat
    ' box from the input area
    messageEditAreaY0 = y0 + d.messages.y() + d.messages.height()
    messageEditAreaX0 = x0

    pie_UniTransBoxFill messageEditAreaX0 + 1, messageEditAreaY0, messageEditAreaX0 + self.width() - 1, y0 + self.height() - 1, pal_RGBA(0, 0, 100, 80)

    iV_Line messageEditAreaX0 + 1, messageEditAreaY0, messageEditAreaX0 + self.width() - 1, messageEditAreaY0, psPalette(WZCOL_FORM_DARK)
End Sub

Public Sub ChatBoxWidget_initializeMessages(ByVal self As WIDGET, ByVal preserveOldChat As Boolean)
    Dim i As Long, message As RoomMessage
    If preserveOldChat Then
        For i = 0 To persistentMessageLocalStorage_size - 1
            message = persistentMessageLocalStorage(i)
            ChatBoxWidget_displayMessage self, message
        Next
    Else
        Erase persistentMessageLocalStorage
        persistentMessageLocalStorage_size = 0
    End If
End Sub

Private Sub ChatBoxWidget_displayMessage(ByVal self As WIDGET, message As RoomMessage)
    Dim d As ChatBoxWidget, paragraph As WIDGET, playerNameWidget As WIDGET, specSender As Boolean
    Dim msgText As String, senderCopy As Long, c As Long
    Set d = ChatBoxWidget_(self)
    Set paragraph = Paragraph_new()
    paragraph.setGeometry 0, 0, ScrollableListWidget_calculateListViewWidth(d.messages), 0

    Select Case message.typ
    Case RoomMessageSystem
        Paragraph_setFontColour paragraph, psPalette(WZCOL_CONS_TEXT_SYSTEM)
        Paragraph_addText paragraph, message.text

    Case RoomMessagePlayer
        If Not (message.sender >= 0) Then ASSERT_FAILURE "message.sender.get() != nullptr", "src/multiint.cpp:4459", "displayMessage", "Null message sender?"
        Paragraph_setFont paragraph, font_small
        Paragraph_setFontColour paragraph, pal_RGBA(&HC0, &HC0, &HC0, &HFF)
        Paragraph_addText paragraph, formatLocalDateTime("%H:%M", message.time_)

        Set playerNameWidget = ChatBoxPlayerNameWidget_new(message.sender)
        Paragraph_addWidget paragraph, playerNameWidget, ChatBoxPlayerNameWidget_aboveBase(playerNameWidget)

        Paragraph_setFont paragraph, font_regular
        Paragraph_setShadeColour paragraph, pal_RGBA(0, 0, 0, 0)
        specSender = NetPlay_players_isSpectator(message.sender) And Not (message.sender = NetPlay_hostPlayer)
        If Not specSender Then c = psPalette(WZCOL_WHITE) Else c = psPalette(WZCOL_TEXT_MEDIUM)
        Paragraph_setFontColour paragraph, c
        Paragraph_addText paragraph, " " & message.text

    Case Else ' RoomMessageNotify
        Paragraph_setFontColour paragraph, psPalette(WZCOL_YELLOW)
        Paragraph_addText paragraph, message.text
    End Select

    msgText = message.text
    If message.typ = RoomMessagePlayer Then senderCopy = message.sender Else senderCopy = -1
    Paragraph_addOnClickHandler paragraph, WzFunc_new(FN_chatBox_paragraphContextMenu, msgText, senderCopy)

    ScrollableListWidget_addItem d.messages, paragraph
End Sub

' (VB) sender: the player index (-1: nullptr; the PlayerReference's detachment comes with the joining, cluster J)
Private Function ChatBoxWidget_createParagraphContextualMenuPopoverForm(ByVal self As WIDGET, ByVal textToCopy As String, ByVal sender As Long) As WIDGET
    Dim weakChatBoxWidget As Long, buttons() As WIDGET, nButtons As Long, senderIdx As Long, currentlyMuted As Boolean
    Dim muteString As String, totalButtonHeight As Long, maxButtonWidth As Long, itemsList As WIDGET, pad As Padding, i As Long
    Const itemSpacing As Long = 4
    ' create all the buttons / option rows
    weakChatBoxWidget = ChatBoxWidget_(self).weakSelf

    ReDim buttons(0 To 1)
    Set buttons(nButtons) = chatBox_createOptionsButton(weakChatBoxWidget, "Copy Text to Clipboard", WzFunc_new(FN_chatBox_copyText, weakChatBoxWidget, textToCopy))
    nButtons = nButtons + 1
    If sender >= 0 Then
        senderIdx = sender
        If senderIdx <> selectedPlayer Then
            currentlyMuted = ingame.muteChat(senderIdx)
            If Not currentlyMuted Then
                muteString = va("Mute Player: %s", NetPlay_players_name(senderIdx))
            Else
                muteString = va("Unmute Player: %s", NetPlay_players_name(senderIdx))
            End If
            Set buttons(nButtons) = chatBox_createOptionsButton(weakChatBoxWidget, muteString, WzFunc_new(FN_chatBox_mutePlayer, sender, currentlyMuted))
            nButtons = nButtons + 1
        End If
    End If

    ' determine required height for all buttons
    For i = 0 To nButtons - 1
        totalButtonHeight = totalButtonHeight + buttons(i).height()
        If buttons(i).width() > maxButtonWidth Then maxButtonWidth = buttons(i).width()
    Next

    Set itemsList = ScrollableListWidget_make()
    ScrollableListWidget_setBackgroundColor itemsList, psPalette(WZCOL_MENU_BACKGROUND)
    pad.top = 3: pad.right_ = 4: pad.bottom = 3: pad.left_ = 4
    ScrollableListWidget_setPadding itemsList, pad
    ScrollableListWidget_setItemSpacing itemsList, itemSpacing
    itemsList.setGeometry itemsList.x(), itemsList.y(), maxButtonWidth + 8, totalButtonHeight + (nButtons * itemSpacing) + 6
    For i = 0 To nButtons - 1
        ScrollableListWidget_addItem itemsList, buttons(i)
    Next

    Set ChatBoxWidget_createParagraphContextualMenuPopoverForm = itemsList
End Function

' createParagraphContextualMenuPopoverForm's createOptionsButton lambda
Private Function chatBox_createOptionsButton(ByVal weakChatBoxWidget As Long, ByVal text As String, ByVal onClickFunc As WzFunc) As WIDGET
    Dim button As WIDGET, cache As PopoverMenuButtonDisplayCache, minButtonWidthForText As Long
    Set button = W_BUTTON_new()
    button.setString text
    W_BUTTON_(button).FontID = font_regular
    Set button.displayFunction = WzFunc_new(FN_PopoverMenuButtonDisplayFunc)
    Set cache = New PopoverMenuButtonDisplayCache
    Set button.pUserData = cache
    cache.text.setText text, W_BUTTON_(button).FontID
    ' (setOnDelete deletes the cache: it goes with the widget)
    minButtonWidthForText = cache.text.width()
    button.setGeometry 0, 0, minButtonWidthForText + MENU_BUTTONS_PADDING, cache.text.lineSize() + 4

    ' On click, perform the designated onClickHandler and close the popover form / overlay
    W_BUTTON_addOnClickHandler button, WzFunc_new(FN_chatBox_menuButton_onClick, weakChatBoxWidget, onClickFunc)
    Set chatBox_createOptionsButton = button
End Function

Private Sub ChatBoxWidget_displayParagraphContextualMenu(ByVal self As WIDGET, ByVal textToCopy As String, ByVal sender As Long)
    Dim d As ChatBoxWidget, lockedScreen As W_SCREEN, newRootFrm As WIDGET, psWeakOptionsOverlayScreen As Long
    Dim psWeakChatBoxWidget As Long, optionsPopOver As WIDGET, popOverX0 As Long, popOverY0 As Long
    Set d = ChatBoxWidget_(self)
    Set lockedScreen = self.screenPointer()
    If Not (Not lockedScreen Is Nothing) Then ASSERT_FAILURE "lockedScreen != nullptr", "src/multiint.cpp:4591", "displayParagraphContextualMenu", "The ChatBoxWidget does not have an associated screen pointer?"

    ' Initialize the options overlay screen
    Set d.optionsOverlayScreen = W_SCREEN_make()
    Set newRootFrm = W_FULLSCREENOVERLAY_CLICKFORM_make()
    Set newRootFrm.displayFunction = WzFunc_new(FN_displayChildDropShadows)
    psWeakOptionsOverlayScreen = ObjPtr(d.optionsOverlayScreen)
    psWeakChatBoxWidget = d.weakSelf
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_chatBox_optionsOverlay_onClicked, psWeakOptionsOverlayScreen, psWeakChatBoxWidget)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onCancelPressed = W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc
    d.optionsOverlayScreen.psForm.attach newRootFrm

    ' Create the pop-over form
    Set optionsPopOver = ChatBoxWidget_createParagraphContextualMenuPopoverForm(self, textToCopy, sender)
    newRootFrm.attach optionsPopOver

    ' Position the pop-over form - use current mouse position
    popOverX0 = MIN_(mouseX(), screenWidth - optionsPopOver.width())
    popOverY0 = MIN_(mouseY(), screenHeight - optionsPopOver.height())
    optionsPopOver.move popOverX0, popOverY0

    widgRegisterOverlayScreenOnTopOfScreen d.optionsOverlayScreen, lockedScreen
End Sub

'--------------------------------------------------------------------------
' class ChatBoxPlayerNameWidget: public WIDGET (player: std::shared_ptr<PlayerReference>, the player's index:
' (*player)->x is NetPlay.players[player].x)

Public Sub ChatBoxPlayerNameWidget_construct(ByVal self As WIDGET, ByVal player As Long)
    WIDGET_construct self, WIDG_UNSPECIFIED_TYPE
    self.kind = K_ChatBoxPlayerNameWidget ' (VB) virtual calls from here on reach ChatBoxPlayerNameWidget's overrides
    With ChatBoxPlayerNameWidget_(self)
        .player = player
        .font = font_regular
        Set .cachedText = WzText_new("", .font)
        .horizontalPadding = 3
        .leftMargin = 3
        .ptsAboveBase = 0
    End With
    ChatBoxPlayerNameWidget_updateLayout self
    self.setTransparentToMouse True
End Sub

Public Function ChatBoxPlayerNameWidget_new(ByVal player As Long) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_ChatBoxPlayerNameWidget)
    ChatBoxPlayerNameWidget_construct w, player
    Set ChatBoxPlayerNameWidget_new = w
End Function

Public Sub ChatBoxPlayerNameWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As ChatBoxPlayerNameWidget, left_ As Long, top As Long, marginLeft As Long, textX As Long, textY As Long
    Dim isSpectator As Boolean, bckColor As Long
    Set d = ChatBoxPlayerNameWidget_(self)
    left_ = xOffset + self.x()
    top = yOffset + self.y()
    marginLeft = left_ + d.leftMargin
    textX = marginLeft + d.horizontalPadding
    textY = top - d.cachedText.aboveBase()
    isSpectator = NetPlay_players_isSpectator(d.player)
    bckColor = pal_RGBA(0, 0, 0, 70)
    If Not isSpectator Then
        bckColor = pal_GetTeamColour(NetPlay_players_colour(d.player))
    End If
    pie_UniTransBoxFill marginLeft, top, left_ + self.width(), top + self.height(), bckColor
    d.cachedText.renderOutlined textX, textY, psPalette(WZCOL_WHITE), pal_RGBA(0, 0, 0, 128)
End Sub

Public Sub ChatBoxPlayerNameWidget_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    If ChatBoxPlayerNameWidget_(self).layoutName <> NetPlay_players_name(ChatBoxPlayerNameWidget_(self).player) Then
        ChatBoxPlayerNameWidget_updateLayout self
    End If
    ' (cachedText.tick(): the WzCachedText's cache expiry; the WzText is kept)
End Sub

Public Function ChatBoxPlayerNameWidget_aboveBase(ByVal self As WIDGET) As Long
    ChatBoxPlayerNameWidget_aboveBase = ChatBoxPlayerNameWidget_(self).ptsAboveBase
End Function

Private Sub ChatBoxPlayerNameWidget_updateLayout(ByVal self As WIDGET)
    Dim d As ChatBoxPlayerNameWidget
    Set d = ChatBoxPlayerNameWidget_(self)
    d.layoutName = NetPlay_players_name(d.player)
    Set d.cachedText = WzText_new(d.layoutName, d.font)
    d.ptsAboveBase = d.cachedText.aboveBase()
    self.setGeometry self.x(), self.y(), d.cachedText.width() + d.leftMargin + 2 * d.horizontalPadding, d.cachedText.lineSize()
End Sub

Public Sub addChatBox(Optional ByVal preserveOldChat As Boolean = False)
    Dim desiredSendMode As ChatBoxSendMode, psExistingChatBoxWidget As WIDGET, parent As WIDGET, chatBox As WIDGET
    Dim modListMessage As String
    If Not widgGetFromID(psWScreen, FRONTEND_TOPFORM) Is Nothing Then
        widgDelete psWScreen, FRONTEND_TOPFORM
    End If

    If headlessGameMode() Then
        Exit Sub
    End If

    desiredSendMode = ChatBoxSendMode_HOSTMSG_ONLY
    If selectedPlayer < MAX_CONNECTED_PLAYERS Then
        If ingame.hostChatPermissions(selectedPlayer) Then desiredSendMode = ChatBoxSendMode_ENABLED
    End If

    Set psExistingChatBoxWidget = widgGetFromID(psWScreen, MULTIOP_CHATBOX)
    If Not psExistingChatBoxWidget Is Nothing Then
        ' chat is already up
        ' update send mode if needed
        If WIDGET_isKindOf(psExistingChatBoxWidget, K_ChatBoxWidget) Then
            ChatBoxWidget_setSendMode psExistingChatBoxWidget, desiredSendMode
        End If
        Exit Sub
    End If

    Set parent = widgGetFromID(psWScreen, FRONTEND_BACKDROP)
    Set chatBox = ChatBoxWidget_make()
    parent.attach chatBox
    chatBox.setCalcLayout WzFunc_new(FN_chatBox_calcLayout)
    ChatBoxWidget_setSendMode chatBox, desiredSendMode
    ChatBoxWidget_initializeMessages chatBox, preserveOldChat

    addSideText FRONTEND_SIDETEXT4, MULTIOP_CHATBOXX - 6, MULTIOP_CHATBOXY, "CHAT"

    If Len(getModList()) <> 0 Then
        modListMessage = "Mod: "
        modListMessage = modListMessage & getModList()
        displayRoomSystemMessage modListMessage
    End If
End Sub

'//////////////////////////////////////////////////////////////////////////
Public Sub stopJoining(ByVal parent As WzTitleUI, ByVal errorResult As LOBBY_ERROR_TYPES)
    Dim origInGameSide As Boolean, needsLevReload As Boolean, i As Long
    updateInActualHostedLobby False

    reloadMPConfig ' reload own settings
    cancelOrDismissVoteNotifications
    ' (the lambda: the tags that start with SLOTTYPE_TAG_PREFIX)
    cancelOrDismissNotificationIfTagHasPrefix SLOTTYPE_TAG_PREFIX

    debug_ 4695, LOG_NET, "stopJoining", "player %u (Host is %s) stopping.", selectedPlayer, IIf(NetPlay_isHost, "true", "false")

    pie_LoadBackDrop SCREEN_RANDOMBDROP

    If NetPlay_isHost Then ' cancel a hosted game.
        ' annouce we are leaving...
        debug_ 4702, LOG_NET, "stopJoining", "Host is quitting game..."
        NETend NETbeginEncode(NETbroadcastQueue(), NET_HOST_DROPPED)
        sendLeavingMsg                                  ' say goodbye
        NETclose                                        ' quit running game.
        ActivityManager_hostLobbyQuit errorResult
        changeTitleUI wzTitleUICurrent                  ' refresh options screen.
        ingame.localJoiningInProgress = False
        Exit Sub
    ElseIf ingame.localJoiningInProgress Then           ' cancel a joined game.
        origInGameSide = ingame.side

        debug_ 4716, LOG_NET, "stopJoining", "Canceling game..."
        sendLeavingMsg                                  ' say goodbye
        NETclose                                        ' quit running game.

        NET_clearDownloadingWZFiles
        ingame.localJoiningInProgress = False           ' reset local flags
        ingame.localOptionsReceived = False

        If origInGameSide = InGameSide_MULTIPLAYER_CLIENT Then
            saveMultiOptionPrefValues sPlayer, selectedPlayer ' persist any changes to multioption preferences

            ingame.side = InGameSide_HOST_OR_SINGLEPLAYER ' must reset to default so that rebuildSearchPath properly rebuilds local mods paths
            needsLevReload = game.modHashCount <> 0
            Erase game.modHashes ' must clear this so that when search paths are reloaded we don't reload mods downloaded for the last game
            game.modHashCount = 0
            rebuildSearchPath mod_multiplay, True
            If needsLevReload Then
                ' Clear & reload level list
                levShutDown
                levInitialise
                pal_Init ' Update palettes.
                If Not buildMapList(False) Then
                    debug_ 4740, LOG_FATAL, "stopJoining", "Failed to rebuild map / level list?"
                End If
            End If
        End If

        ' joining and host was transferred.
        If origInGameSide = InGameSide_MULTIPLAYER_CLIENT And NetPlay_isHost Then
            NetPlay_isHost = False
        End If

        ActivityManager_joinedLobbyQuit errorResult
        If Not parent Is Nothing Then
            changeTitleUI parent
        Else
            changeTitleMode MULTI
        End If

        selectedPlayer = 0
        realSelectedPlayer = 0
        Exit Sub
    End If
    debug_ 4765, LOG_NET, "stopJoining", "We have stopped joining."
    ActivityManager_joinedLobbyQuit errorResult
    NETremRedirects
    changeTitleUI parent
    selectedPlayer = 0
    realSelectedPlayer = 0
    For i = 0 To NetPlay_players_size() - 1
        NetPlay_players_resetAll i
    Next
    NetPlay_players_resize MAX_CONNECTED_PLAYERS
    For i = 0 To NetPlay_scriptSetPlayerDataStrings_size() - 1
        NetPlay_scriptSetPlayerDataStrings_clear i
    Next
End Sub

Public Sub resetPlayerPositions()
    Dim i As Long
    ' Reset players' positions or bad things could happen to scavenger slot

    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        NetPlay_players_position(i) = i
        NetPlay_players_team(i) = i
        NETBroadcastPlayerInfo i
    Next
End Sub

Public Function repositionHumanSlots() As Long
    Dim pos As Long, i As Long
    pos = 0

    ' First, put human players at the top
    For i = 0 To MAX_PLAYERS - 1
        If isHumanPlayer(i) Then
            ' Skip the scavenger slot
            If game.mapHasScavengers And pos = scavengerSlot() Then
                pos = pos + 1
            End If
            NetPlay_players_position(i) = pos
            NETBroadcastPlayerInfo i
            pos = pos + 1
        End If
    Next

    repositionHumanSlots = pos
End Function

Public Sub updateMapSettings(ByVal mapData As LEVEL_DATASET)
    If Not (Not mapData Is Nothing) Then ASSERT_FAILURE "mapData != nullptr", "src/multiint.cpp:4819", "updateMapSettings", "Invalid mapData?": Exit Sub
    game.map = mapData.pName
    game.hash = levGetFileHash(mapData)
    game.maxPlayers = mapData.players
    game.isMapMod = CheckForMod(mapData.realFileName)
    game.isRandom = CheckForRandom(mapData.realFileName, mapData.apDataFiles(0))
End Sub

Private Sub randomizeLimit(ByVal name_ As String)
    Dim stat As Long
    stat = getStructStatFromName(name_)
    If Not (stat >= 0) Then ASSERT_FAILURE "stat >= 0", "src/multiint.cpp:5239", "randomizeLimit", "Invalid struct stat name: %s", name_: Exit Sub
    If rand() Mod 2 = 0 Then
        asStructureStats(stat).upgrade(0).limit = asStructureStats(stat).base.limit
    Else
        asStructureStats(stat).upgrade(0).limit = 0
    End If
End Sub

' Generate random options
Public Sub multiLobbyRandomizeOptions()
    Dim pos As Long, i As Long
    If ingame.side <> InGameSide_HOST_OR_SINGLEPLAYER Then Exit Sub ' RUN_ONLY_ON_SIDE(InGameSide::HOST_OR_SINGLEPLAYER)

    If NetPlay_bComms And NetPlay_isHost And Not canChangeMapOrRandomize() Then
        Exit Sub
    End If

    resetPlayerPositions

    ' Reset and randomize player positions, also to guard
    ' against case where in the previous map some players
    ' took slots, which aren't available anymore

    pos = repositionHumanSlots()

    ' Fill with AIs if places remain
    For i = 0 To current_numplayers - 1
        If Not isHumanPlayer(i) And Not (game.mapHasScavengers And NetPlay_players_position(i) = scavengerSlot()) Then
            ' Skip the scavenger slot
            If game.mapHasScavengers And pos = scavengerSlot() Then
                pos = pos + 1
            End If
            NetPlay_players_position(i) = pos
            NETBroadcastPlayerInfo i
            pos = pos + 1
        End If
    Next

    ' Randomize positions
    For i = 0 To current_numplayers - 1
        SendPositionRequest i, NetPlay_players_position(rand() Mod current_numplayers)
    Next

    ' Structure limits are simply 0 or max, only for NO_TANK, NO_CYBORG, NO_VTOL, NO_UPLINK, NO_LASSAT.
    If Not bLimiterLoaded Or numStructureStats = 0 Then
        initLoadingScreen True
        If resLoad("wrf/limiter_data.wrf", 503) Then
            bLimiterLoaded = True
        End If
        closeLoadingScreen
    End If
    resetLimits
    For i = 0 To limitIcons_size() - 2 ' skip last item, MPFLAGS_FORCELIMITS
        randomizeLimit limitIcons(i).stat
    Next
    If rand() Mod 2 = 0 Then
        ingame.flags = ingame.flags Or MPFLAGS_FORCELIMITS
    Else
        ingame.flags = ingame.flags And Not MPFLAGS_FORCELIMITS
    End If
    createLimitSet
    updateStructureDisabledFlags

    ' Game options
    If Not locked.scavengers And game.mapHasScavengers Then
        game.scavengers = umod(rand(), SCAV_TYPE_MAX + 1)
    End If

    If Not locked.alliances Then
        game.alliance = umod(rand(), ALLIANCE_TYPE_MAX + 1)
    End If
    If Not locked.power Then
        game.power = umod(rand(), POWER_SETTING_MAX + 1)
    End If
    If Not locked.bases Then
        game.base = umod(rand(), CAMP_TYPE_MAX + 1)
    End If

    game.techLevel = umod(rand(), TECH_LEVEL_MAX) + TECH_LEVEL_MIN

    If NetPlay_isHost Then
        resetReadyStatus True
    End If

    refreshMultiplayerOptionsTitleUIIfActive

    NETsetLobbyConfigFlagsFields game.alliance, game.techLevel, game.power, game.base
End Sub

' (VB) static_cast<uint32_t>(rand()) % m: rand() is never negative
Private Function umod(ByVal v As Long, ByVal m As Long) As Long
    umod = v Mod m
End Function

Public Sub handleAutoReadyRequest()
    Dim desiredReadyState As Boolean, haveData As Boolean
    If NetPlay_isHost Then
        Exit Sub ' no-op for host (currently)
    End If

    If GetGameMode() = GS_NORMAL Then
        Exit Sub ' don't bother sending anything - the game is starting / started...
    End If

    desiredReadyState = NetPlay_players_ready(selectedPlayer)

    ' Spectators should automatically become ready as soon as necessary files are downloaded
    ' and not-ready when files remain to be downloaded
    If NetPlay_players_isSpectator(selectedPlayer) Then
        haveData = NET_getDownloadingWzFiles_empty()
        desiredReadyState = haveData
    End If

    If desiredReadyState <> NetPlay_players_ready(selectedPlayer) Then
        ' Automatically set ready status
        sendReadyRequest selectedPlayer, desiredReadyState
    End If
End Sub

Public Sub multiClearHostRequestMoveToPlayer(ByVal playerIdx As Long)
    If Not ULT(playerIdx, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/multiint.cpp:5632", "multiClearHostRequestMoveToPlayer", "Invalid playerIdx: %u", playerIdx: Exit Sub
    If Not NetPlay_isHost Then Exit Sub
    bHostRequestedMoveToPlayers(playerIdx) = False
End Sub

' (class WzHostLobbyOperationsInterface : public HostLobbyOperationsInterface, the lobby slash commands' host
' operations, and its cmdInterface: the lobby commands, cluster J)

Public Function getBoundedMinAutostartPlayerCount() As Long
    Dim minAutoStartPlayerCount As Long, numAIPlayers As Long, maxPlayerSlots As Long, j As Long
    minAutoStartPlayerCount = min_autostart_player_count()
    If minAutoStartPlayerCount <= 0 Then
        ' there is no minimum configured
        getBoundedMinAutostartPlayerCount = -1
        Exit Function
    End If
    For j = 0 To MAX_PLAYERS - 1
        If (NetPlay_players_allocated(j) Or NetPlay_players_ai(j) = AI_OPEN) And NetPlay_players_position(j) < game.maxPlayers Then
            maxPlayerSlots = maxPlayerSlots + 1
        ElseIf Not NetPlay_players_allocated(j) And NetPlay_players_ai(j) >= 0 Then
            numAIPlayers = numAIPlayers + 1
        End If
    Next
    If minAutoStartPlayerCount > maxPlayerSlots Then
        debug_ 5879, LOG_WARNING, "getBoundedMinAutostartPlayerCount", "startplayers (%d) exceeds maxPlayerSlots (%d) - using the latter; (game.maxPlayers: %u, # ai players: %d)", minAutoStartPlayerCount, maxPlayerSlots, game.maxPlayers, numAIPlayers
        minAutoStartPlayerCount = maxPlayerSlots
    End If

    getBoundedMinAutostartPlayerCount = minAutoStartPlayerCount
End Function

Public Sub calcBackdropLayoutForMultiplayerOptionsTitleUI(ByVal psWidget As WIDGET)
    Dim height As Long
    height = screenHeight - 80
    CLIP height, HIDDEN_FRONTEND_HEIGHT, HIDDEN_FRONTEND_WIDTH

    psWidget.setGeometry _
        ((screenWidth - HIDDEN_FRONTEND_WIDTH) \ 2), _
        ((screenHeight - height) \ 2), _
        HIDDEN_FRONTEND_WIDTH - 1, _
        height
End Sub

'--------------------------------------------------------------------------
' class WzPlayerSelectionChangePositionRow : public W_BUTTON

Public Sub WzPlayerSelectionChangePositionRow_construct(ByVal self As WIDGET, ByVal targetPlayerIdx As Long)
    W_BUTTON_construct self
    self.kind = K_WzPlayerSelectionChangePositionRow ' (VB) virtual calls from here on reach WzPlayerSelectionChangePositionRow's overrides
    WzPlayerSelectionChangePositionRow_(self).targetPlayerIdx = targetPlayerIdx
    ' (DisplayPositionCache cache)
    Set WzPlayerSelectionChangePositionRow_(self).cache = New WzText
End Sub

Public Function WzPlayerSelectionChangePositionRow_make(ByVal switcherPlayerIdx As Long, ByVal targetPlayerIdx As Long, ByVal parent As WzMultiplayerOptionsTitleUI) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzPlayerSelectionChangePositionRow)
    WzPlayerSelectionChangePositionRow_construct widget, targetPlayerIdx
    widget.setTip "Click to change to this slot"

    W_BUTTON_addOnClickHandler widget, WzFunc_new(FN_changePositionRow_onClick, switcherPlayerIdx, parent.selfPtr())

    Set WzPlayerSelectionChangePositionRow_make = widget
End Function

Public Sub WzPlayerSelectionChangePositionRow_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, text As String, cache As WzText
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    drawBlueBox x0, y0, self.width(), self.height()
    text = va("Click to take player slot %u", NetPlay_players_position(WzPlayerSelectionChangePositionRow_(self).targetPlayerIdx))
    Set cache = WzPlayerSelectionChangePositionRow_(self).cache
    cache.setText text, font_regular
    cache.render x0 + 10, y0 + 22, psPalette(WZCOL_FORM_TEXT)
End Sub

'--------------------------------------------------------------------------
' class WzPlayerIndexSwapPositionRow : public W_BUTTON

Public Sub WzPlayerIndexSwapPositionRow_construct(ByVal self As WIDGET, ByVal targetPlayerIdx As Long)
    W_BUTTON_construct self
    self.kind = K_WzPlayerIndexSwapPositionRow ' (VB) virtual calls from here on reach WzPlayerIndexSwapPositionRow's overrides
    WzPlayerIndexSwapPositionRow_(self).targetPlayerIdx = targetPlayerIdx
End Sub

Public Function WzPlayerIndexSwapPositionRow_make(ByVal switcherPlayerIdx As Long, ByVal targetPlayerIdx As Long, ByVal parent As WzMultiplayerOptionsTitleUI) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzPlayerIndexSwapPositionRow)
    WzPlayerIndexSwapPositionRow_construct widget, targetPlayerIdx
    If targetPlayerIdx <> NetPlay_hostPlayer Then
        widget.setTip "Click to swap player to this slot"

        W_BUTTON_addOnClickHandler widget, WzFunc_new(FN_swapPositionRow_onClick, switcherPlayerIdx, parent.selfPtr())
    Else
        widget.setTip "Cannot swap with host"
        widget.setState WBUT_DISABLE
    End If

    Set WzPlayerIndexSwapPositionRow_make = widget
End Function

Public Sub WzPlayerIndexSwapPositionRow_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, isDisabled As Boolean
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    isDisabled = (self.getState() And WBUT_DISABLE) <> 0

    If isDisabled Then
        pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), pal_RGBA(0, 0, 0, 125)
        Exit Sub
    End If

    iV_Box x0, y0, x0 + self.width(), y0 + self.height(), psPalette(WZCOL_MENU_BORDER)
End Sub

' WzPlayerSelectPositionRowFactory::make (the factory: an enum, its make a Select)
Private Function WzPlayerSelectPositionRowFactory_make(ByVal rowFactory As WzPlayerSelectPositionRowFactory, ByVal switcherPlayerIdx As Long, ByVal targetPlayerIdx As Long, ByVal parent As WzMultiplayerOptionsTitleUI) As WIDGET
    Select Case rowFactory
    Case WzPlayerSelectionChangePositionRowFactory
        Set WzPlayerSelectPositionRowFactory_make = WzPlayerSelectionChangePositionRow_make(switcherPlayerIdx, targetPlayerIdx, parent)
    Case WzPlayerIndexSwapPositionRowRactory
        Set WzPlayerSelectPositionRowFactory_make = WzPlayerIndexSwapPositionRow_make(switcherPlayerIdx, targetPlayerIdx, parent)
    End Select
End Function

'--------------------------------------------------------------------------
' class WzPositionChooser : public WIDGET

Public Sub WzPositionChooser_construct(ByVal self As WIDGET, ByVal switcherPlayerIdx As Long, ByVal parentTitleUI As WzMultiplayerOptionsTitleUI)
    WIDGET_construct self
    self.kind = K_WzPositionChooser ' (VB) virtual calls from here on reach WzPositionChooser's overrides
    WzPositionChooser_(self).switcherPlayerIdx = switcherPlayerIdx
    WzPositionChooser_(self).psWeakTitleUI = parentTitleUI.selfPtr()
    WzPositionChooser_(self).weakSelf = WeakWidget_register(self)      ' (VB) its std::weak_ptr<WzPositionChooser>s
End Sub

' (VB) ~WzPositionChooser: the weak pointers to it expire
Public Sub WzPositionChooser_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzPositionChooser_(self).weakSelf
End Sub

Public Function WzPositionChooser_make(ByVal switcherPlayerIdx As Long, ByVal rowFactory As WzPlayerSelectPositionRowFactory, ByVal parent As WzMultiplayerOptionsTitleUI) As WIDGET
    Dim widget As WIDGET, i As Long, playerPositionRow As WIDGET
    Set widget = WIDGET_alloc(K_WzPositionChooser)
    WzPositionChooser_construct widget, switcherPlayerIdx, parent

    ' Create WzPlayerSelectPositionRow all player positions
    For i = 0 To game.maxPlayers - 1
        Set playerPositionRow = WzPlayerSelectPositionRowFactory_make(rowFactory, switcherPlayerIdx, i, parent)
        widget.attach playerPositionRow
        WzPositionChooser_(widget).positionSelectionRows_push_back playerPositionRow, i
    Next
    WzPositionChooser_positionRows widget

    ' Position each in the appropriate spot
    Set WzPositionChooser_make = widget
End Function

Public Sub WzPositionChooser_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim psWeakTitleUICopy As Long
    psWeakTitleUICopy = WzPositionChooser_(self).psWeakTitleUI
    widgScheduleTask WzFunc_new(FN_positionChooser_clicked_task, psWeakTitleUICopy)
End Sub

Public Sub WzPositionChooser_geometryChanged(ByVal self As WIDGET)
    WzPositionChooser_positionRows self
End Sub

Public Sub WzPositionChooser_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzPositionChooser_positionRows self
End Sub

Public Sub WzPositionChooser_positionRows(ByVal self As WIDGET)
    Dim d As WzPositionChooser, strongTitleUI As WzMultiplayerOptionsTitleUI, validPlayerIdxRowsToDisplay() As Boolean
    Dim i As Long, row_widget As WIDGET, row_targetPlayerIdx As Long, newWidth As Long
    Set d = WzPositionChooser_(self)
    Set strongTitleUI = titleUI_lock(d.psWeakTitleUI)
    If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "strongTitleUI != nullptr", "src/multiint.cpp:1944", "positionRows", "No valid TitleUI pointer": Exit Sub
    validPlayerIdxTargetsForPlayerPositionMove d.switcherPlayerIdx, validPlayerIdxRowsToDisplay
    For i = 0 To d.positionSelectionRows_size() - 1
        Set row_widget = d.positionSelectionRows_widget(i)
        row_targetPlayerIdx = d.positionSelectionRows_targetPlayerIdx(i)
        If Not validPlayerIdxRowsToDisplay(row_targetPlayerIdx) Then
            row_widget.hide
        Else
            newWidth = self.width()
            row_widget.setGeometry 0, MAX_(strongTitleUI.playerRowY0(row_targetPlayerIdx), 0), newWidth, MULTIOP_PLAYERHEIGHT
            row_widget.show
        End If
    Next
End Sub

'/////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////
' Drawing functions

Public Sub displayChatEdit(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, borderColor As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    borderColor = psPalette(WZCOL_MENU_SEPARATOR)
    If Not W_EDITBOX_isEditing(psWidget) Then
        borderColor = pielightWithAlpha(borderColor, pielightAlpha(borderColor) \ 2)
    End If
    iV_Box x, y, x + psWidget.width(), y + psWidget.height(), borderColor
End Sub

' ////////////////////////////////////////////////////////////////////////////

Public Sub displaySpectatorAddButton(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    drawBlueBox_SpectatorOnly x, y, psWidget.width(), psWidget.height()
    iV_DrawImage FrontImages, IMAGE_SPECTATOR, x + 2, y + 1
End Sub

Public Sub displayAi(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim cache As WzText, x As Long, y As Long, j As Long, displayText As String, textColor As Long, id As Long
    ' Any widget using displayAi must have its pUserData initialized to a (DisplayAICache*)
    ' (assert(psWidget->pUserData != nullptr); its wzText: the pUserData itself)
    Set cache = psWidget.pUserData

    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    j = psWidget.UserData

    drawBlueBox x, y, psWidget.width(), psWidget.height()
    textColor = psPalette(WZCOL_FORM_TEXT)

    If j >= 0 Then
        If j < getAIData_size() Then
            displayText = getAIData_name(j)
        Else
            debug_ 7021, LOG_ERROR, "displayAi", "Invalid aidata index: %d; (num AIs: %zu)", j, getAIData_size()
        End If
    Else
        id = psWidget.id
        Select Case id
        Case MULTIOP_AI_CLOSED
            displayText = "Closed"
        Case MULTIOP_AI_OPEN
            displayText = "Open"
        Case MULTIOP_AI_SPECTATOR
            displayText = "Spectator"
        Case Else
            debug_ 7039, LOG_ERROR, "displayAi", "Unexpected button id: %d", id
        End Select
    End If
    cache.setText displayText, font_regular
    cache.render x + 10, y + 22, textColor
End Sub

Public Function difficultyIcon(ByVal difficulty As Long) As Long
    Select Case difficulty
    Case AIDifficulty_EASY: difficultyIcon = IMAGE_EASY
    Case AIDifficulty_MEDIUM: difficultyIcon = IMAGE_MEDIUM
    Case AIDifficulty_HARD: difficultyIcon = IMAGE_HARD
    Case AIDifficulty_INSANE: difficultyIcon = IMAGE_INSANE
    Case Else: difficultyIcon = IMAGE_NO   '/ what??
    End Select
End Function

Private Sub displayDifficulty(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim cache As WzText, x As Long, y As Long, j As Long
    ' Any widget using displayDifficulty must have its pUserData initialized to a (DisplayDifficultyCache*)
    ' (its wzDifficultyText: the pUserData itself)
    Set cache = psWidget.pUserData

    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    j = psWidget.UserData

    If Not (j < getDifficultyListCount()) Then ASSERT_FAILURE "j < difficultyList.size()", "src/multiint.cpp:7069", "displayDifficulty", "Bad difficulty found: %d", j: Exit Sub
    drawBlueBox x, y, psWidget.width(), psWidget.height()
    iV_DrawImage FrontImages, difficultyIcon(j), x + 5, y + 5
    cache.setText difficultyList_at(j), font_regular
    cache.render x + 42, y + 22, psPalette(WZCOL_FORM_TEXT)
End Sub

' ////////////////////////////////////////////////////////////////////////////
' Display blue box
Public Sub intDisplayFeBox(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, w As Long, h As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    w = psWidget.width()
    h = psWidget.height()

    drawBlueBox x, y, w, h
End Sub

Private Sub drawBlueBox_Spectator(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim backgroundClr As Long
    ' Match drawBlueBox behavior
    x = x - 1
    y = y - 1
    w = w + 2
    h = h + 2
    backgroundClr = psPalette(WZCOL_FORM_DARK)
    backgroundClr = pielightWithAlpha(backgroundClr, 200)
    pie_BoxFill x, y, x + w, y + h, psPalette(WZCOL_MENU_BORDER)
    pie_UniTransBoxFill x + 1, y + 1, x + w - 1, y + h - 1, backgroundClr
End Sub

Private Sub drawBlueBox_SpectatorOnly(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim backgroundClr As Long
    ' Match drawBlueBox behavior
    x = x - 1
    y = y - 1
    w = w + 2
    h = h + 2
    backgroundClr = psPalette(WZCOL_FORM_DARK)
    backgroundClr = pielightWithAlpha(backgroundClr, 175)
'   PIELIGHT borderClr = pal_RGBA(30, 30, 30, 120);
'   pie_UniTransBoxFill(x, y, x + w, y + h, borderClr);
'   pie_UniTransBoxFill(x + 1, y + 1, x + w - 1, y + h - 1, backgroundClr);
    pie_UniTransBoxFill x, y, x + w, y + h, backgroundClr
End Sub

' ////////////////////////////////////////////////////////////////////////////
' Display (not-blue) box for spectators
Public Sub intDisplayFeBox_Spectator(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, w As Long, h As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    w = psWidget.width()
    h = psWidget.height()

    drawBlueBox_Spectator x, y, w, h
End Sub

Public Sub intDisplayFeBox_SpectatorOnly(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, w As Long, h As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    w = psWidget.width()
    h = psWidget.height()

    drawBlueBox_SpectatorOnly x, y, w, h
End Sub

' ////////////////////////////////////////////////////////////////////////////
' Display edit box
Public Sub displayMultiEditBox(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, colour As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    drawBlueBox x, y, psWidget.width(), psWidget.height()

    If (W_EDITBOX_(psWidget).state And WEDBS_DISABLE) <> 0 Then             ' disabled
        colour = pal_RGBA(FILLRED, FILLGREEN, FILLBLUE, FILLTRANS)
        pie_UniTransBoxFill x, y, x + psWidget.width() + psWidget.height() - 1, y + psWidget.height(), colour
    End If
End Sub

Public Function mpwidgetGetFrontHighlightImage(image As AtlasImage) As AtlasImage
    If AtlasImage_isNull(image) Then
        mpwidgetGetFrontHighlightImage = AtlasImage_()
        Exit Function
    End If
    Select Case AtlasImage_width(image.images, image.id)
    Case 30: mpwidgetGetFrontHighlightImage = AtlasImage_(FrontImages, IMAGE_HI34)
    Case 60: mpwidgetGetFrontHighlightImage = AtlasImage_(FrontImages, IMAGE_HI64)
    Case 19: mpwidgetGetFrontHighlightImage = AtlasImage_(FrontImages, IMAGE_HI23)
    Case 27: mpwidgetGetFrontHighlightImage = AtlasImage_(FrontImages, IMAGE_HI31)
    Case 35: mpwidgetGetFrontHighlightImage = AtlasImage_(FrontImages, IMAGE_HI39)
    Case 37: mpwidgetGetFrontHighlightImage = AtlasImage_(FrontImages, IMAGE_HI41)
    Case 56: mpwidgetGetFrontHighlightImage = AtlasImage_(FrontImages, IMAGE_HI56)
    Case Else: mpwidgetGetFrontHighlightImage = AtlasImage_()
    End Select
End Function

'--------------------------------------------------------------------------
' struct WzMultiButton : public W_BUTTON

' WzMultiButton() : W_BUTTON() {} (and the members' initialisers)
Public Sub WzMultiButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzMultiButton ' (VB) virtual calls from here on reach WzMultiButton's overrides
    With WzMultiButton_(self)
        .alpha = 255
        .downStateMask = WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK
        .greyStateMask = WBUT_DISABLE
        .has_drawBlueBorder = False
    End With
End Sub

Public Function WzMultiButton_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_WzMultiButton)
    WzMultiButton_construct w
    Set WzMultiButton_new = w
End Function

Public Sub WzMultiButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzMultiButton, x0 As Long, y0 As Long, hiToUse As AtlasImage, highlight As Boolean, down As Boolean, grey As Boolean
    Dim toDraw(0 To 2) As AtlasImage, numToDraw As Long, n As Long, tcImage As AtlasImage, imNormal As AtlasImage
    Dim f As Long, mix As Long, drawBlueBorder_or As Boolean
    Const scale_ As Long = 4000
    Set d = WzMultiButton_(self)
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    hiToUse = AtlasImage_(0, 0)
    imNormal = d.imNormal

    ' FIXME: This seems to be a way to conserve space, so you can use a
    ' transparent icon with these edit boxes.
    ' hack for multieditbox
    drawBlueBorder_or = d.has_drawBlueBorder And d.drawBlueBorder      ' drawBlueBorder.value_or(false)
    If imNormal.id = IMAGE_EDIT_MAP Or imNormal.id = IMAGE_EDIT_GAME Or imNormal.id = IMAGE_EDIT_PLAYER _
       Or imNormal.id = IMAGE_LOCK_BLUE Or imNormal.id = IMAGE_UNLOCK_BLUE Or drawBlueBorder_or Then
        If Not d.has_drawBlueBorder Or d.drawBlueBorder Then ' if drawBlueBox is explicitly set to false, don't draw the box
            drawBlueBox x0 - 2, y0 - 2, self.height(), self.height()    ' box on end.
        End If
    End If

    ' evaluate auto-frame
    highlight = (self.getState() And WBUT_HIGHLIGHT) <> 0

    ' evaluate auto-frame
    If d.doHighlight = 1 And highlight Then
        hiToUse = mpwidgetGetFrontHighlightImage(imNormal)
    End If

    down = (self.getState() And d.downStateMask) <> 0
    grey = (self.getState() And d.greyStateMask) <> 0

    numToDraw = 0

    ' now display
    toDraw(numToDraw) = imNormal
    numToDraw = numToDraw + 1

    ' hilights etc..
    If down Then
        toDraw(numToDraw) = d.imDown
        numToDraw = numToDraw + 1
    End If
    If highlight And Not grey And hiToUse.images <> 0 Then
        toDraw(numToDraw) = hiToUse
        numToDraw = numToDraw + 1
    End If

    For n = 0 To numToDraw - 1
        tcImage = AtlasImage_(toDraw(n).images, toDraw(n).id + 1)
        If d.tc = MAX_PLAYERS Then
            iV_DrawImageImage toDraw(n).images, toDraw(n).id, x0, y0, d.alpha
        ElseIf d.tc = MAX_PLAYERS + 1 Then
            f = realTime Mod scale_
            mix = pal_RGBA(128 + iSinR(65536 * f \ scale_ + 65536 * 0 \ 3, 127), _
                           128 + iSinR(65536 * f \ scale_ + 65536 * 1 \ 3, 127), _
                           128 + iSinR(65536 * f \ scale_ + 65536 * 2 \ 3, 127), _
                           255)
            iV_DrawImageTc toDraw(n).images, toDraw(n).id, tcImage.images, tcImage.id, x0, y0, mix
        Else
            iV_DrawImageTc toDraw(n).images, toDraw(n).id, tcImage.images, tcImage.id, x0, y0, pal_GetTeamColour(d.tc)
        End If
    Next

    If grey Then
        ' disabled, render something over it!
        iV_TransBoxFill x0, y0, x0 + self.width(), y0 + self.height()
    End If
End Sub

'/////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////
' common widgets

Public Function addMultiEditBox(ByVal formid As Long, ByVal id As Long, ByVal x As Long, ByVal y As Long, ByVal tip As String, ByVal tipres As String, ByVal icon As Long, ByVal iconhi As Long, ByVal iconid As Long, ByVal disabled As Boolean) As WIDGET
    Dim sEdInit As W_EDBINIT, editBox As WIDGET, multiBut As WIDGET
    ' editbox
    sEdInit.formID = formid
    sEdInit.id = id
    sEdInit.x = x                       ' (short)x
    sEdInit.y = y
    sEdInit.width = MULTIOP_EDITBOXW
    sEdInit.height = MULTIOP_EDITBOXH
    sEdInit.pText = tipres
    Set sEdInit.pBoxDisplay = WzFunc_new(FN_displayMultiEditBox)
    Set editBox = widgAddEditBox(psWScreen, sEdInit)
    If editBox Is Nothing Then
        Exit Function
    End If

    Set multiBut = addMultiBut_screen(psWScreen, MULTIOP_OPTIONS, iconid, x + MULTIOP_EDITBOXW + 2, y + 2, MULTIOP_EDITBOXH, MULTIOP_EDITBOXH, tip, icon, iconhi, iconhi)

    If disabled Then
        editBox.setState WEDBS_DISABLE
        multiBut.setState WBUT_DISABLE
    End If

    Set addMultiEditBox = editBox
End Function

Public Sub updateMultiBut(ByVal button As WIDGET, ByVal width As Long, ByVal height As Long, ByVal tipres As String, ByVal norm As Long, ByVal down As Long, ByVal hi As Long, ByVal tc As Long, Optional ByVal alpha As Long = 255)
    button.setGeometry button.x(), button.y(), width, height
    button.setTip tipres          ' (tipres != nullptr) ? std::string(tipres) : std::string()
    WzMultiButton_(button).imNormal = AtlasImage_(FrontImages, norm)
    WzMultiButton_(button).imDown = AtlasImage_(FrontImages, down)
    WzMultiButton_(button).doHighlight = hi
    WzMultiButton_(button).tc = tc
    WzMultiButton_(button).alpha = alpha
End Sub

Public Function makeMultiBut(ByVal id As Long, ByVal width As Long, ByVal height As Long, ByVal tipres As String, ByVal norm As Long, ByVal down As Long, ByVal hi As Long, ByVal tc As Long, Optional ByVal alpha As Long = 255) As WIDGET
    Dim button As WIDGET
    Set button = WzMultiButton_new()
    button.id = id
    updateMultiBut button, width, height, tipres, norm, down, hi, tc, alpha
    Set makeMultiBut = button
End Function

'/////////////////////////////////////////////////////////////////////////////////////////
Public Function addMultiBut(ByVal parent As WIDGET, ByVal id As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal tipres As String, ByVal norm As Long, ByVal down As Long, ByVal hi As Long, Optional ByVal tc As Long = MAX_PLAYERS, Optional ByVal alpha As Long = 255) As WIDGET
    Dim button As WIDGET, existingWidget As WIDGET
    Set existingWidget = widgFormGetFromID(parent, id)
    If Not existingWidget Is Nothing Then
        If WIDGET_isKindOf(existingWidget, K_WzMultiButton) Then Set button = existingWidget
    End If
    If button Is Nothing Then
        Set button = WzMultiButton_new()
        parent.attach button
        button.id = id
    End If
    button.move x, y
    updateMultiBut button, width, height, tipres, norm, down, hi, tc, alpha
    Set addMultiBut = button
End Function

' std::shared_ptr<W_BUTTON> addMultiBut(const std::shared_ptr<W_SCREEN> &screen, UDWORD formid, ...)
Public Function addMultiBut_screen(ByVal screen As W_SCREEN, ByVal formid As Long, ByVal id As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal tipres As String, ByVal norm As Long, ByVal down As Long, ByVal hi As Long, Optional ByVal tc As Long = MAX_PLAYERS, Optional ByVal alpha As Long = 255) As WIDGET
    Dim pWidget As WIDGET
    Set pWidget = widgGetFromID(screen, formid)
    If Not (Not pWidget Is Nothing) Then ASSERT_FAILURE "pWidget != nullptr", "src/multiint.cpp:7329", "addMultiBut", "Unable to find formID: %u", formid
    Set addMultiBut_screen = addMultiBut(pWidget, id, x, y, width, height, tipres, norm, down, hi, tc, alpha)
End Function

Public Function multiplayLacksEnoughPlayersToAutostart() As Boolean
    Dim allowStart As Boolean, minAutoStartPlayerCount As Long, playersPresent As Long, j As Long
    If Not NetPlay_isHost Then
        ' host does not share min autostart player count with clients,
        ' so just assume it's disabled (and host will start/autostart when appropriate)
        multiplayLacksEnoughPlayersToAutostart = False
        Exit Function
    End If
    allowStart = True
    minAutoStartPlayerCount = getBoundedMinAutostartPlayerCount()
    If minAutoStartPlayerCount > 0 Then
        playersPresent = 0
        For j = 0 To MAX_PLAYERS - 1
            If NetPlay_players_allocated(j) Then
                playersPresent = playersPresent + 1
            End If
        Next
        If playersPresent < minAutoStartPlayerCount Then
            allowStart = False
        End If
    End If
    multiplayLacksEnoughPlayersToAutostart = Not allowStart
End Function

' Returns true if the multiplayer game can start (i.e. all players are ready)
Public Function multiplayPlayersReady() As Boolean
    Dim bReady As Boolean, numReadyPlayers As Long, player As Long
    bReady = True
    numReadyPlayers = 0

    For player = 0 To NetPlay_players_size() - 1
        ' check if this human player is ready, ignore AIs
        If NetPlay_players_allocated(player) Then
            If (Not NetPlay_players_ready(player) Or ingame.PingTimes(player) >= PING_LIMIT) Then
                bReady = False
            Else
                numReadyPlayers = numReadyPlayers + 1
            End If
        ElseIf NetPlay_players_ai(player) >= 0 Then
            numReadyPlayers = numReadyPlayers + 1
        End If
    Next

    multiplayPlayersReady = bReady And (numReadyPlayers > 1) And (allPlayersOnSameTeam(-1) = -1)
End Function

Public Function multiplayIsStartingGame() As Boolean
    multiplayIsStartingGame = bInActualHostedLobby And multiplayPlayersReady() And Not multiplayLacksEnoughPlayersToAutostart()
End Function

Public Sub sendRoomNotifyMessage(ByVal text As String)
    Dim message As NetworkTextMessage
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:7443", "sendRoomNotifyMessage", "Host only routine detected for client!": Exit Sub
    message = NetworkTextMessage_new(NOTIFY_MESSAGE, text)
    displayRoomSystemMessage text
    NetworkTextMessage_enqueue message, NETbroadcastQueue()
End Sub

Public Sub sendRoomSystemMessageToSingleReceiver(ByVal text As String, ByVal receiver As Long, Optional ByVal skipLocalDisplay As Boolean = False)
    Dim message As NetworkTextMessage
    If Not isHumanPlayer(receiver) Then ASSERT_FAILURE "isHumanPlayer(receiver)", "src/multiint.cpp:7451", "sendRoomSystemMessageToSingleReceiver", "Invalid receiver: %u", receiver: Exit Sub
    message = NetworkTextMessage_new(SYSTEM_MESSAGE, text)
    If Not skipLocalDisplay Or receiver = selectedPlayer Then
        displayRoomSystemMessage text
    End If
    NetworkTextMessage_enqueue message, NETnetQueue(receiver)
End Sub

Public Sub sendRoomChatMessage(ByVal text As String, Optional ByVal skipLocalDisplay As Boolean = False)
    Dim mutedUntil As Currency, currentTime As Currency, message As NetworkTextMessage, m As RoomMessage
    If NetPlay_bComms Then
        If playerSpamMutedUntil(selectedPlayer, mutedUntil) Then
            QueryPerformanceCounter currentTime
            If perfFrequency = 0 Then QueryPerformanceFrequency perfFrequency
            addConsoleMessage "You have sent too many messages in the last few seconds. Please wait and try again.", DEFAULT_JUSTIFY, INFO_MESSAGE, False, CLng(Fix(CDbl(mutedUntil - currentTime) * 1000# / CDbl(perfFrequency)))
            Exit Sub
        End If
        recordPlayerMessageSent selectedPlayer
    End If

    message = NetworkTextMessage_new(selectedPlayer, text)
    If Not skipLocalDisplay Then
        m = RoomMessage_player(selectedPlayer, text)
        displayRoomMessage m
    End If
    If NetPlay_isHost Then
        ' Always allow the host to execute lobby slash commands
        If processChatLobbySlashCommands(message) Then
            ' consume the message
            Exit Sub
        End If
    End If
    If Left$(message.text, LOBBY_COMMAND_PREFIX_LENGTH + 7) = LOBBY_COMMAND_PREFIX & "hostmsg" Then
        ' send the message to the host only
        NetworkTextMessage_enqueue message, NETnetQueue(NetPlay_hostPlayer)
    Else
        NetworkTextMessage_enqueue message, NETbroadcastQueue()
    End If
End Sub

Public Function numSlotsToBeDisplayed() As Long
    Dim n As Long, i As Long
    If Not (MAX_CONNECTED_PLAYERS <= NetPlay_players_size()) Then ASSERT_FAILURE "static_cast<size_t>(((11 + 2) + 10)) <= NetPlay.players.size()", "src/multiint.cpp:7501", "numSlotsToBeDisplayed", "Insufficient array size: %zu versus %zu", NetPlay_players_size(), MAX_CONNECTED_PLAYERS
    n = 0
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If isSpectatorOnlySlot(i) _
           And ((i >= NetPlay_players_size()) Or Not (NetPlay_players_isSpectator(i) And NetPlay_players_ai(i) = AI_OPEN)) Then
            ' the only slots displayable beyond maxPlayers are spectator-only slots that are enabled
        Else
            n = n + 1
        End If
    Next
    numSlotsToBeDisplayed = n
End Function

Public Function spectatorSlotsSupported() As Boolean
    spectatorSlotsSupported = NetPlay_bComms
End Function

' NOTE: Pass in the index in the NetPlay.players array
Public Function isSpectatorOnlySlot(ByVal playerIdx As Long) As Boolean
    If Not ULT(playerIdx, NetPlay_players_size()) Then ASSERT_FAILURE "playerIdx < NetPlay.players.size()", "src/multiint.cpp:7524", "isSpectatorOnlySlot", "Invalid playerIdx: %u", playerIdx: isSpectatorOnlySlot = False: Exit Function
    isSpectatorOnlySlot = playerIdx >= MAX_PLAYERS Or NetPlay_players_position(playerIdx) >= game.maxPlayers
End Function

'--------------------------------------------------------------------------
' (VB) the lambdas (docs/CONVENTIONS.md, "Widgets"); each logs under its containing function
Public Function multiint_lobby_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, psParent As WIDGET, pStrongPtr As WzMultiplayerOptionsTitleUI, strongTitleUI As WzMultiplayerOptionsTitleUI
    Dim button As WIDGET, player As Long, id As Long, x0 As Long, y0 As Long, width As Long, showHostButton As Boolean
    Dim parentHeight As Long, hostButtonVerticalSpace As Long, maxOptionsListHeight As Long, msg As String, playerName As String
    Dim strongParentForm As WIDGET, str_ As String, psChatBoxWidget As WIDGET, psMessages As WIDGET, m As RoomMessage
    Dim quickChatWidth As Long, chatBoxScreenPosX As Long, chatBoxScreenWidth As Long, idealChatOptionsWidth As Long
    Dim maxWidth As Long, w As Long, x As Long, idealChatOptionsHeight As Long, maxHeight As Long, lastIdealChatOptionsHeight As Long
    Dim psOverlayScreen As W_SCREEN, chatBoxWidget As WIDGET, psStrongPositionChooser As WIDGET, height As Long
    Select Case f.id

    Case FN_intDisplayFeBox
        intDisplayFeBox args(0), args(1), args(2)
    Case FN_intDisplayFeBox_Spectator
        intDisplayFeBox_Spectator args(0), args(1), args(2)
    Case FN_intDisplayFeBox_SpectatorOnly
        intDisplayFeBox_SpectatorOnly args(0), args(1), args(2)
    Case FN_displayAi
        displayAi args(0), args(1), args(2)
    Case FN_displayDifficulty
        displayDifficulty args(0), args(1), args(2)
    Case FN_displayChatEdit
        displayChatEdit args(0), args(1), args(2)
    Case FN_displaySpectatorAddButton
        displaySpectatorAddButton args(0), args(1), args(2)
    Case FN_displayMultiEditBox
        displayMultiEditBox args(0), args(1), args(2)
    Case FN_backdrop_calcLayout
        calcBackdropLayoutForMultiplayerOptionsTitleUI args(0)

    Case FN_addInlineChooserBlueForm_calcLayout      ' [x, y, psWeakParent](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = f.cap(2)
        If Not psParent Is Nothing Then
            psWidget.move psParent.screenPosX() + f.cap(0), psParent.screenPosY() + f.cap(1)
        End If

    Case FN_optionsForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry MULTIOP_OPTIONSX, MULTIOP_OPTIONSY, MULTIOP_OPTIONSW, MULTIOP_OPTIONSH

    Case FN_startHosting_onClick                     ' [psWeakTitleUI](W_BUTTON&)
        Set strongTitleUI = titleUI_lock(f.cap(0))
        If Not strongTitleUI Is Nothing Then
            strongTitleUI.startHost
        End If

    Case FN_multiLobbyOptionsForm_calcLayout         ' [psWeakTitleUI](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/multiint.cpp:1261", "addGameOptions", "No parent?": Exit Function
        Set strongTitleUI = titleUI_lock(f.cap(0))
        If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "strongTitleUI != nullptr", "src/multiint.cpp:1263", "addGameOptions", "No titleUI?": Exit Function
        x0 = 1
        y0 = 42
        width = psParent.width() - 2
        If Not strongTitleUI.getStartHostingButton() Is Nothing Then showHostButton = strongTitleUI.getStartHostingButton().visible() Else showHostButton = False
        parentHeight = psParent.height()
        If showHostButton Then hostButtonVerticalSpace = (parentHeight - strongTitleUI.getStartHostingButton().y()) Else hostButtonVerticalSpace = 0
        maxOptionsListHeight = parentHeight - y0 - hostButtonVerticalSpace
        psWidget.setGeometry x0, y0, width, MIN_(maxOptionsListHeight, psWidget.idealHeight())

    Case FN_rightSideChooser_calcLayout              ' [psWeakParent](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = f.cap(0)
        If Not psParent Is Nothing Then
            psWidget.setGeometry psParent.screenPosX() + MULTIOP_PLAYERSX + f.cap(1), psParent.screenPosY() + MULTIOP_PLAYERSY, MULTIOP_PLAYERSW, MULTIOP_PLAYERSH
        End If

    Case FN_difficultyRow_onClick                    ' [psWeakTitleUI, difficultyIdx, player](W_BUTTON& clickedButton)
        Set pStrongPtr = titleUI_lock(f.cap(0))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:1482", "openDifficultyChooser", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function
        player = f.cap(2)
        NetPlay_players_difficulty(player) = difficultyValue_at(f.cap(1))
        NETBroadcastPlayerInfo player
        resetReadyStatus False, isBlindSimpleLobby(game.blindMode)
        widgScheduleTask WzFunc_new(FN_task_closeDifficultyChooser, pStrongPtr)

    Case FN_task_closeDifficultyChooser              ' [pStrongPtr]
        Set pStrongPtr = f.cap(0)
        pStrongPtr.closeDifficultyChooser
        pStrongPtr.addPlayerBox True

    Case FN_aiOpenClose_onClick                      ' [psWeakTitleUI, player](W_BUTTON& clickedButton)
        Set button = args(0)
        Set pStrongPtr = titleUI_lock(f.cap(0))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:1579", "openAiChooser", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function
        player = f.cap(1)

        If Not NetPlay_players_allocated(player) Then
            Select Case button.id
            Case MULTIOP_AI_CLOSED
                changeAI player, AI_CLOSED
                NetPlay_players_isSpectator(player) = False
            Case MULTIOP_AI_OPEN
                changeAI player, AI_OPEN
                NetPlay_players_isSpectator(player) = False
            Case MULTIOP_AI_SPECTATOR
                ' set slot to open
                changeAI player, AI_OPEN
                ' but also a spectator
                NetPlay_players_isSpectator(player) = True
            Case Else
                debug_ 1600, LOG_ERROR, "openAiChooser", "Unexpected button id"
                Exit Function
            End Select

            ' common code
            NetPlay_players_difficulty(player) = AIDifficulty_DISABLED ' disable AI for this slot
            NETBroadcastPlayerInfo player
            resetReadyStatus False, isBlindSimpleLobby(game.blindMode)
        Else
            debug_ 1612, LOG_INFO, "openAiChooser", "Player joined slot %d while chooser was open", player
        End If
        widgScheduleTask WzFunc_new(FN_task_closeAiChooser, pStrongPtr)
        ActivityManager_updateMultiplayGameData

    Case FN_task_closeAiChooser                      ' [pStrongPtr]
        Set pStrongPtr = f.cap(0)
        pStrongPtr.closeAiChooser
        pStrongPtr.addPlayerBox True

    Case FN_aiList_calcLayout                        ' [aiListStartXPos, aiListStartYPos, aiListEntryWidth, aiListHeight]
        Set psWidget = args(0)
        psWidget.setGeometry f.cap(0), f.cap(1), f.cap(2), f.cap(3)

    Case FN_aiRow_onClick                            ' [psWeakTitleUI, aiIdx, player](W_BUTTON& clickedButton)
        Set pStrongPtr = titleUI_lock(f.cap(0))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:1700", "openAiChooser", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function
        player = f.cap(2)
        If Not NetPlay_players_allocated(player) Then
            NetPlay_players_isSpectator(player) = False
            changeAI player, f.cap(1)
            setPlayerName player, getAIName(player)
            ' Difficulty is preserved when switching open AI slots.
            If NetPlay_players_difficulty(player) = AIDifficulty_DISABLED Then
                NetPlay_players_difficulty(player) = AIDifficulty_DEFAULT
            End If
            NETBroadcastPlayerInfo player
            resetReadyStatus False
        Else
            debug_ 1716, LOG_INFO, "openAiChooser", "Player joined slot %d while chooser was open", player
        End If
        widgScheduleTask WzFunc_new(FN_task_closeAiChooser, pStrongPtr)
        ActivityManager_updateMultiplayGameData

    Case FN_initPUserData_DisplayAICache             ' []() -> void * { return new DisplayAICache(); }
        Set multiint_lobby_c_fn = WzText_new("", font_regular)

    Case FN_changePositionRow_onClick                ' [switcherPlayerIdx, titleUI](W_BUTTON& button)
        Set button = args(0)
        If Not (WIDGET_isKindOf(button, K_WzPlayerSelectionChangePositionRow)) Then ASSERT_FAILURE "selectPositionRow != nullptr", "src/multiint.cpp:1758", "make", "Wrong widget type": Exit Function
        Set strongTitleUI = titleUI_lock(f.cap(1))
        If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "strongTitleUI != nullptr", "src/multiint.cpp:1760", "make", "Title UI is gone?": Exit Function
        ' Switch player
        SendPositionRequest f.cap(0), NetPlay_players_position(WzPlayerSelectionChangePositionRow_(button).targetPlayerIdx)
        widgScheduleTask WzFunc_new(FN_task_closePositionChooser, strongTitleUI)

    Case FN_task_closePositionChooser                ' [strongTitleUI]
        Set strongTitleUI = f.cap(0)
        strongTitleUI.closePositionChooser
        strongTitleUI.updatePlayers

    Case FN_swapPositionRow_onClick                  ' [switcherPlayerIdx, titleUI](W_BUTTON& button)
        Set button = args(0)
        If Not (WIDGET_isKindOf(button, K_WzPlayerIndexSwapPositionRow)) Then ASSERT_FAILURE "selectPositionRow != nullptr", "src/multiint.cpp:1821", "make", "Wrong widget type": Exit Function
        Set strongTitleUI = titleUI_lock(f.cap(1))
        If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "strongTitleUI != nullptr", "src/multiint.cpp:1823", "make", "Title UI is gone?": Exit Function

        playerName = getPlayerName(f.cap(0), True)

        NETmoveSpectatorToPlayerSlot f.cap(0), WzPlayerIndexSwapPositionRow_(button).targetPlayerIdx, True
        resetReadyStatus True, isBlindSimpleLobby(game.blindMode)      'reset and send notification to all clients
        widgScheduleTask WzFunc_new(FN_task_closePlayerSlotSwapChooser, strongTitleUI)
        msg = va("Spectator %s has moved to Players", playerName)
        sendRoomSystemMessage msg

    Case FN_task_closePlayerSlotSwapChooser          ' [strongTitleUI]
        Set strongTitleUI = f.cap(0)
        strongTitleUI.closePlayerSlotSwapChooser
        strongTitleUI.updatePlayers

    Case FN_positionChooser_clicked_task             ' [psWeakTitleUICopy]
        Set strongTitleUI = titleUI_lock(f.cap(0))
        If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "psTitleUI != nullptr", "src/multiint.cpp:1925", "clicked", "Title UI is null": Exit Function
        strongTitleUI.closeAllChoosers ' this also removes the overlay screen
        strongTitleUI.updatePlayers

    Case FN_team_onClick                             ' [player, psWeakTitleUI](W_BUTTON &button)
        Set button = args(0)
        id = button.id
        Set pStrongPtr = titleUI_lock(f.cap(1))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:2087", "openTeamChooser", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function
        player = f.cap(0)

        'clicked on a team
        ' STATIC_ASSERT(MULTIOP_TEAMCHOOSER + MAX_PLAYERS - 1 <= MULTIOP_TEAMCHOOSER_END);
        If id >= MULTIOP_TEAMCHOOSER And id <= MULTIOP_TEAMCHOOSER + MAX_PLAYERS - 1 Then
            If Not (id >= MULTIOP_TEAMCHOOSER And (id - MULTIOP_TEAMCHOOSER) < MAX_PLAYERS) Then ASSERT_FAILURE "id >= 102810 && (id - 102810) < 11", "src/multiint.cpp:2093", "openTeamChooser", "processMultiopWidgets: wrong id - MULTIOP_TEAMCHOOSER value (%d)", id - MULTIOP_TEAMCHOOSER

            resetReadyStatus False, isBlindSimpleLobby(game.blindMode)     ' will reset only locally if not a host

            SendTeamRequest player, (id - MULTIOP_TEAMCHOOSER) And &HFF&

            debug_ 2100, LOG_WZ, "openTeamChooser", "Changed team for player %d to %d", player, NetPlay_players_team(player)

            widgScheduleTask WzFunc_new(FN_task_closeTeamChooser_addPlayerBox, pStrongPtr)
        End If

    Case FN_task_closeTeamChooser_addPlayerBox       ' [pStrongPtr]
        Set pStrongPtr = f.cap(0)
        pStrongPtr.closeTeamChooser
        ' restore player list
        pStrongPtr.addPlayerBox True

    Case FN_kick_onClick                             ' [player, psWeakTitleUI](W_BUTTON &button)
        Set pStrongPtr = titleUI_lock(f.cap(1))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:2135", "openTeamChooser", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function
        player = f.cap(0)

        msg = va("The host has kicked %s from the game!", getPlayerName(player, True))
        kickPlayer player, "The host has kicked you from the game.", ERROR_KICKED, False
        sendRoomSystemMessage msg
        resetReadyStatus True       'reset and send notification to all clients
        widgScheduleTask WzFunc_new(FN_task_closeTeamChooser, pStrongPtr)

    Case FN_task_closeTeamChooser                    ' [pStrongPtr]
        Set pStrongPtr = f.cap(0)
        pStrongPtr.closeTeamChooser

    Case FN_ban_onClick                              ' [player, psWeakTitleUI](W_BUTTON &button)
        Set pStrongPtr = titleUI_lock(f.cap(1))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:2155", "openTeamChooser", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function
        player = f.cap(0)

        msg = va("The host has banned %s from the game!", getPlayerName(player, True))
        kickPlayer player, "The host has banned you from the game.", ERROR_KICKED, True
        sendRoomSystemMessage msg
        resetReadyStatus True       'reset and send notification to all clients
        widgScheduleTask WzFunc_new(FN_task_closeTeamChooser, pStrongPtr)

    Case FN_specHost_onClick                         ' [player, psWeakTitleUI](W_BUTTON &button)
        Set pStrongPtr = titleUI_lock(f.cap(1))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:2187", "openTeamChooser", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function
        player = f.cap(0)

        playerName = getPlayerName(player, True)

        If Not NETmovePlayerToSpectatorOnlySlot(player, True) Then
            msg = va("Failed to move %s to Spectators", playerName)
            sendRoomSystemMessageToSingleReceiver msg, selectedPlayer
            Exit Function
        End If

        msg = va("The host has moved %s to Spectators!", playerName)
        sendRoomSystemMessage msg
        resetReadyStatus True       'reset and send notification to all clients
        widgScheduleTask WzFunc_new(FN_task_closeTeamChooser, pStrongPtr)

    Case FN_specClient_onClick                       ' [player, psWeakTitleUI](W_BUTTON &button)
        Set pStrongPtr = titleUI_lock(f.cap(1))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:2210", "openTeamChooser", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function

        ' Send net message to host asking to be a spectator, as host is responsible for these changes
        SendPlayerSlotTypeRequest f.cap(0), True
        widgScheduleTask WzFunc_new(FN_task_closeTeamChooser, pStrongPtr)

    Case FN_askToPlay_onClick                        ' [player, psWeakTitleUI](W_BUTTON &button)
        Set pStrongPtr = titleUI_lock(f.cap(1))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:" & f.cap(2), "openTeamChooser", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function

        ' Ask the spectator if they are okay with a move from spectator -> player? (host) / Send net message to
        ' host asking to be a player, as host is responsible for these changes (client)
        SendPlayerSlotTypeRequest f.cap(0), False
        widgScheduleTask WzFunc_new(FN_task_closeTeamChooser, pStrongPtr)

    Case FN_colour_onClick                           ' [player, psWeakTitleUI](W_BUTTON &button)
        Set button = args(0)
        id = button.id
        Set pStrongPtr = titleUI_lock(f.cap(1))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:2297", "openColourChooser", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function

        If id >= MULTIOP_COLCHOOSER And id < MULTIOP_COLCHOOSER + MAX_PLAYERS - 1 Then  ' chose a new colour.
            resetReadyStatus False, True        ' will reset only locally if not a host
            SendColourRequest f.cap(0), id - MULTIOP_COLCHOOSER
            widgScheduleTask WzFunc_new(FN_task_closeColourChooser, pStrongPtr)
        End If

    Case FN_faction_onClick                          ' [player, psWeakTitleUI](W_BUTTON &button)
        Set button = args(0)
        id = button.id
        Set pStrongPtr = titleUI_lock(f.cap(1))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:2334", "openColourChooser", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function

        ' STATIC_ASSERT(MULTIOP_FACCHOOSER + NUM_FACTIONS - 1 <= MULTIOP_FACCHOOSER_END);
        If id >= MULTIOP_FACCHOOSER And id <= MULTIOP_FACCHOOSER + NUM_FACTIONS - 1 Then
            resetReadyStatus False, True
            SendFactionRequest f.cap(0), (id - MULTIOP_FACCHOOSER) And &HFF&
            widgScheduleTask WzFunc_new(FN_task_closeColourChooser, pStrongPtr)
        End If

    Case FN_task_closeColourChooser                  ' [pStrongPtr]
        Set pStrongPtr = f.cap(0)
        pStrongPtr.closeColourChooser
        pStrongPtr.addPlayerBox True

    Case FN_inlineOverlay_onClicked                  ' [psWeakTitleUI]()
        widgScheduleTask WzFunc_new(FN_inlineOverlay_task, f.cap(0))

    Case FN_inlineOverlay_task                       ' [psWeakTitleUI]
        Set strongTitleUI = titleUI_lock(f.cap(0))
        If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "psTitleUI != nullptr", "src/multiint.cpp:6838", "start", "Title UI is null": Exit Function
        strongTitleUI.closeAllChoosers ' this also removes the overlay screen
        strongTitleUI.updatePlayers

    Case FN_playerSlotSwapChooser_calcLayout         ' openPlayerSlotSwapChooser's calcLayouts
        Set psWidget = args(0)
        Select Case f.cap(0)
        Case SWAPLAYOUT_AIFORM                       ' [psWeakParent, additionalHeightForSwappingRow]
            Set psParent = f.cap(1)
            If Not psParent Is Nothing Then
                psWidget.setGeometry psParent.screenPosX() + MULTIOP_PLAYERSX, psParent.screenPosY() + MULTIOP_PLAYERSY, MULTIOP_PLAYERSW, MULTIOP_PLAYERSH + f.cap(2)
            End If
        Case SWAPLAYOUT_POSITIONCHOOSER              ' []
            psWidget.setGeometry 0, 0, MULTIOP_PLAYERSW, MULTIOP_PLAYERSH
        Case SWAPLAYOUT_CONTEXTFORM                  ' [psWeakPositionChooser, swapContextFormHeight, swapContextFormMargin]
            Set psStrongPositionChooser = WeakWidget_lock(f.cap(1))
            If Not psStrongPositionChooser Is Nothing Then
                psWidget.setGeometry 0, psStrongPositionChooser.y() + psStrongPositionChooser.height() + f.cap(3), psStrongPositionChooser.width(), f.cap(2)
            End If
        Case SWAPLAYOUT_PLAYERROW                    ' [swapContextFormPadding, textHeight]
            psWidget.setGeometry PLAYERBOX_X0, f.cap(1) + f.cap(2) + 4, MULTIOP_PLAYERWIDTH, MULTIOP_PLAYERHEIGHT
        Case SWAPLAYOUT_ROWLABEL                     ' [swapContextFormPadding, textHeight]
            psWidget.setGeometry PLAYERBOX_X0, f.cap(1), MULTIOP_PLAYERWIDTH, f.cap(2)
        Case SWAPLAYOUT_SWITCHARROW                  ' [swapContextFormPadding, textHeight]
            Set psParent = psWidget.parent()
            If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/multiint.cpp:3558", "openPlayerSlotSwapChooser", "No parent?": Exit Function
            psWidget.setGeometry psParent.width() - PLAYERBOX_X0 - psWidget.width(), f.cap(1), psWidget.width(), f.cap(2)
        Case SWAPLAYOUT_SIDETEXT                     ' [psWeakParent]
            Set psParent = f.cap(1)
            If Not psParent Is Nothing Then
                psWidget.setGeometry psParent.screenPosX() + MULTIOP_PLAYERSX - 3, psParent.screenPosY() + MULTIOP_PLAYERSY, MULTIOP_PLAYERSW, MULTIOP_PLAYERSH
            End If
        End Select

    Case FN_addSpectatorSlots_onClick                ' [psWeakTitleUI](W_BUTTON& clickedButton)
        Set pStrongPtr = titleUI_lock(f.cap(0))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:3772", "addPlayerBox", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function
        widgScheduleTask WzFunc_new(FN_addSpectatorSlots_task, f.cap(0))

    Case FN_addSpectatorSlots_task                   ' [psWeakTitleUI]
        Set pStrongPtr = titleUI_lock(f.cap(0))
        If Not (Not pStrongPtr Is Nothing) Then ASSERT_FAILURE "pStrongPtr.operator bool()", "src/multiint.cpp:3775", "addPlayerBox", "WzMultiplayerOptionsTitleUI no longer exists": Exit Function
        If NETopenNewSpectatorSlot() Then
            pStrongPtr.addPlayerBox True
'           resetReadyStatus(false);
        End If

    Case FN_playersForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry MULTIOP_PLAYERSX, MULTIOP_PLAYERSY, MULTIOP_PLAYERSW, MULTIOP_PLAYERSH

    Case FN_playersTabs_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry PLAYERBOX_X0 - 1, 0, MULTIOP_PLAYERSW - PLAYERBOX_X0, MULTIOP_PLAYERS_TABS_H

    Case FN_blindWaitingRoom_calcLayout
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/multiint.cpp:3653", "addPlayerBox", "No parent?": Exit Function
        x0 = PLAYERBOX_X0 - 1
        y0 = 1
        width = MULTIOP_PLAYERSW - (PLAYERBOX_X0 * 2)
        height = psParent.height() - y0 - 1
        psWidget.setGeometry x0, y0, width, height

    Case FN_chatBox_calcLayout
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not psParent Is Nothing Then
            psWidget.setGeometry MULTIOP_CHATBOXX, MULTIOP_CHATBOXY, MULTIOP_CHATBOXW, psParent.height() - MULTIOP_CHATBOXY
        End If

    Case FN_chatBox_consoleListener                  ' [&](ConsoleMessage const &message): args (sender, text)
        Set chatBoxWidget = WeakWidget_lock(f.cap(0))
        If Not chatBoxWidget Is Nothing Then
            m = buildMessage(args(0), args(1))
            ChatBoxWidget_addMessage chatBoxWidget, m
        End If

    Case FN_chatBox_quickChat_onClick                ' [weakSelf](W_BUTTON& widg)
        widgScheduleTask WzFunc_new(FN_chatBox_openQuickChat_task, f.cap(0), 3980)

    Case FN_chatBox_openQuickChat_task               ' [weakSelf]() (cap(1): the C++ line)
        Set strongParentForm = WeakWidget_lock(f.cap(0))
        If Not (Not strongParentForm Is Nothing) Then ASSERT_FAILURE "strongParentForm != nullptr", "src/multiint.cpp:" & f.cap(1), IIf(f.cap(1) = 4218, "run", "initialize"), "No parent form": Exit Function
        ChatBoxWidget_openQuickChatOverlay strongParentForm

    Case FN_chatBox_options_onClick                  ' [weakSelf](W_BUTTON& widg)
        widgScheduleTask WzFunc_new(FN_chatBox_displayOptions_task, f.cap(0))

    Case FN_chatBox_displayOptions_task              ' [weakSelf]()
        Set strongParentForm = WeakWidget_lock(f.cap(0))
        If Not (Not strongParentForm Is Nothing) Then ASSERT_FAILURE "strongParentForm != nullptr", "src/multiint.cpp:3996", "initialize", "No parent form": Exit Function
        ChatBoxWidget_displayOptionsOverlay strongParentForm, ChatBoxWidget_(strongParentForm).optionsButton

    Case FN_chatBox_onEditReturn                     ' [weakSelf](W_EDITBOX& widg)
        Set psWidget = args(0)
        ' don't send empty lines to other players in the lobby
        str_ = psWidget.getString()
        If Len(str_) = 0 Then
            Exit Function
        End If

        Set strongParentForm = WeakWidget_lock(f.cap(0))
        If Not (Not strongParentForm Is Nothing) Then ASSERT_FAILURE "strongParentForm != nullptr", "src/multiint.cpp:4019", "initialize", "No parent form": Exit Function

        Select Case ChatBoxWidget_(strongParentForm).currentMode
        Case ChatBoxSendMode_DISABLED
            displayRoomSystemMessage "The host has disabled free chat. Please use Quick Chat."
            Exit Function
        Case ChatBoxSendMode_HOSTMSG_ONLY
            ' Always ensure sent message starts with /hostmsg
            If Left$(str_, Len(LOBBY_COMMAND_PREFIX & "hostmsg")) <> LOBBY_COMMAND_PREFIX & "hostmsg" Then
                str_ = LOBBY_COMMAND_PREFIX & "hostmsg " & str_
            End If
        Case ChatBoxSendMode_ENABLED
            ' just send message as-is
        End Select

        sendRoomChatMessage str_, (ChatBoxWidget_(strongParentForm).currentMode <> ChatBoxSendMode_ENABLED)
        psWidget.setString ""

    Case FN_chatBox_onTab                            ' [weakSelf](W_EDITBOX& widg) -> bool
        ' on tab handler
        widgScheduleTask WzFunc_new(FN_chatBox_openQuickChat_task, f.cap(0), 4045)
        multiint_lobby_c_fn = True

    Case FN_chatBox_overlay_onClicked                ' [psWeakChatBoxWidget]()
        widgScheduleTask WzFunc_new(FN_chatBox_closeQuickChat_task, f.cap(0), 4060)

    Case FN_chatBox_quickChatSent                    ' [psWeakChatBoxWidget]() (on quick chat sent: close the overlay)
        widgScheduleTask WzFunc_new(FN_chatBox_closeQuickChat_task, f.cap(0), 4072)

    Case FN_chatBox_closeQuickChat_task              ' [psWeakChatBoxWidget] (cap(1): the C++ line)
        Set psChatBoxWidget = WeakWidget_lock(f.cap(0))
        If Not (Not psChatBoxWidget Is Nothing) Then ASSERT_FAILURE "psChatBoxWidget != nullptr", "src/multiint.cpp:" & f.cap(1), "initialize", "ChatBoxWidget is null": Exit Function
        ChatBoxWidget_closeQuickChatOverlay psChatBoxWidget

    Case FN_chatBox_quickChatForm_calcLayout         ' [psWeakChatBoxWidget](WIDGET* psWidget)
        Set psWidget = args(0)
        Set psChatBoxWidget = WeakWidget_lock(f.cap(0))
        If Not (Not psChatBoxWidget Is Nothing) Then ASSERT_FAILURE "psChatBoxWidget != nullptr", "src/multiint.cpp:4079", "initialize", "ChatBoxWidget is null": Exit Function
        quickChatWidth = psChatBoxWidget.width()
        x0 = psChatBoxWidget.screenPosX()
        y0 = ChatBoxWidget_calculateQuickChatScreenPosY(psChatBoxWidget)
        psWidget.setGeometry x0, y0, quickChatWidth, psWidget.idealHeight()

    Case FN_chatBox_optionsOverlay_onClicked         ' [psWeakOptionsOverlayScreen, psWeakChatBoxWidget]()
        Set psOverlayScreen = W_SCREEN_fromPtr(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            widgRemoveOverlayScreen psOverlayScreen
        End If
        ' Destroy Options overlay / overlay screen
        Set chatBoxWidget = WeakWidget_lock(f.cap(1))
        If Not chatBoxWidget Is Nothing Then
            Set ChatBoxWidget_(chatBoxWidget).optionsOverlayScreen = Nothing
        End If

    Case FN_chatBox_chatOptionsChanged               ' []() (on settings change, refresh the chat box to reflect permission changes)
        addChatBox

    Case FN_chatBox_chatOptionsForm_calcLayout       ' [psWeakChatBoxWidget](WIDGET* psWidget)
        Set psWidget = args(0)
        Set psChatBoxWidget = WeakWidget_lock(f.cap(0))
        If Not (Not psChatBoxWidget Is Nothing) Then ASSERT_FAILURE "psChatBoxWidget != nullptr", "src/multiint.cpp:4299", "displayOptionsOverlay", "ChatBoxWidget is null": Exit Function
        chatBoxScreenPosX = psChatBoxWidget.screenPosX()
        chatBoxScreenWidth = psChatBoxWidget.width()
        idealChatOptionsWidth = psWidget.idealWidth()
        maxWidth = screenWidth - 40
        w = MIN_(idealChatOptionsWidth, maxWidth)
        x0 = chatBoxScreenPosX + (chatBoxScreenWidth - w) \ 2
        ' Try to position it above the chat widget
        idealChatOptionsHeight = psWidget.idealHeight()
        y0 = MAX_(20, psChatBoxWidget.screenPosY() - idealChatOptionsHeight)
        maxHeight = screenHeight - y0 - 20
        psWidget.setGeometry x0, y0, w, MIN_(maxHeight, idealChatOptionsHeight)
        ' idealHeight may have changed due to the width change, so check and set again...
        lastIdealChatOptionsHeight = idealChatOptionsHeight
        idealChatOptionsHeight = psWidget.idealHeight()
        If lastIdealChatOptionsHeight <> idealChatOptionsHeight Then
            y0 = MAX_(20, psChatBoxWidget.screenPosY() - idealChatOptionsHeight)
            psWidget.setGeometry x0, y0, w, MIN_(maxHeight, idealChatOptionsHeight)
        End If

    Case FN_chatBox_paragraphContextMenu             ' [msgText, senderCopy](Paragraph& paragraph, WIDGET_KEY key)
        Set psWidget = args(0)
        ' take advantage of the fact that this widget is the great-grand-child of the ChatBoxWidget
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/multiint.cpp:4488", "displayMessage", "Expected parent?": Exit Function
        Set psMessages = psParent.parent()
        If Not WIDGET_isKindOf(psMessages, K_ScrollableListWidget) Then Set psMessages = Nothing
        If Not (Not psMessages Is Nothing) Then ASSERT_FAILURE "psMessages != nullptr", "src/multiint.cpp:4490", "displayMessage", "Expected grand-parent missing?": Exit Function
        Set psChatBoxWidget = psMessages.parent()
        If Not WIDGET_isKindOf(psChatBoxWidget, K_ChatBoxWidget) Then Set psChatBoxWidget = Nothing
        If Not (Not psChatBoxWidget Is Nothing) Then ASSERT_FAILURE "psChatBoxWidget != nullptr", "src/multiint.cpp:4492", "displayMessage", "Expected great-grand-parent missing?": Exit Function
        ' display contextual menu
        ChatBoxWidget_displayParagraphContextualMenu psChatBoxWidget, f.cap(0), f.cap(1)

    Case FN_chatBox_menuButton_onClick               ' [weakChatBoxWidget, onClickFunc](W_BUTTON& button)
        Set chatBoxWidget = WeakWidget_lock(f.cap(0))
        If Not chatBoxWidget Is Nothing Then
            widgRemoveOverlayScreen ChatBoxWidget_(chatBoxWidget).optionsOverlayScreen
            WzFunc_call f.cap(1), args(0)
            Set ChatBoxWidget_(chatBoxWidget).optionsOverlayScreen = Nothing
        End If

    Case FN_chatBox_copyText                         ' [weakChatBoxWidget, textToCopy](W_BUTTON& button)
        Set chatBoxWidget = WeakWidget_lock(f.cap(0))
        If Not chatBoxWidget Is Nothing Then
            wzSetClipboardText f.cap(1)
        End If

    Case FN_chatBox_mutePlayer                       ' [sender, currentlyMuted](W_BUTTON& button)
        ' (!sender->isDetached(): the player references come with the joining, cluster J)
        setPlayerMuted f.cap(0), Not f.cap(1)
    End Select
End Function
