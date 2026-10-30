VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMultiplayerOptionsTitleUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/multiint.cpp, titleui/multiplayer.h: class WzMultiplayerOptionsTitleUI,
' the multiplayer / skirmish options screen (the lobby). Made by
' WzMultiplayerOptionsTitleUI_new (multiint_c, which keeps multiint.cpp's
' statics and game state; multiint_lobby_c has the screen's free functions,
' widget classes and lambdas). The methods the lambdas call are Public (C++'s
' private members a lambda in the class's scope reaches). A lambda's
' std::weak_ptr to the screen is ObjPtr(Me), which multiint_lobby_c's
' titleUI_lock gives back while this is the current lobby screen.
Option Explicit

Implements WzTitleUI

' lib/netplay/nettypes.h
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS
' multiint.cpp
Private Const PLAYERBOX_X0 As Long = 7

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private psInlineChooserOverlayScreen As W_SCREEN ' std::shared_ptr<W_SCREEN> = nullptr
Private parent As WzTitleUI                     ' std::shared_ptr<WzTitleUI>
Private playerRows() As WIDGET                  ' std::vector<std::shared_ptr<WIDGET>>
Private playerRows_size As Long
Private multiLobbyOptionsForm As WIDGET
Private startHostingButton As WIDGET            ' std::shared_ptr<W_BUTTON>
Private performedFirstStart As Boolean

Private inlineChooserUp As Long                 ' int8_t
Private aiChooserUp As Long                     ' int8_t
Private difficultyChooserUp As Long             ' int8_t
Private has_playerSlotSwapChooserUp As Boolean  ' optional<uint32_t> playerSlotSwapChooserUp: has_value()
Private playerSlotSwapChooserUp As Long         ' its value

' WzMultiplayerOptionsTitleUI::WzMultiplayerOptionsTitleUI(std::shared_ptr<WzTitleUI> parent)
Friend Sub init(ByVal parent_ As WzTitleUI)
    Set parent = parent_
    inlineChooserUp = -1
    aiChooserUp = -1
    difficultyChooserUp = -1
End Sub

' WzMultiplayerOptionsTitleUI::~WzMultiplayerOptionsTitleUI()
Private Sub Class_Terminate()
    closeMultiRequester
    If Not psInlineChooserOverlayScreen Is Nothing Then widgRemoveOverlayScreen psInlineChooserOverlayScreen
    closeAllChoosers
    bInActualHostedLobby = False ' don't bother calling updateInActualHostedLobby here because we're destroying everything
    ' (VB) the std::weak_ptr currentMultiOptionsTitleUI expires
    If currentMultiOptionsTitleUI = ObjPtr(Me) Then currentMultiOptionsTitleUI = 0
End Sub

' (VB) the lambdas' std::weak_ptr<WzMultiplayerOptionsTitleUI>
Public Function selfPtr() As Long
    selfPtr = ObjPtr(Me)
End Function

' need to check for side effects.
Public Sub addGameOptions()
    Dim psExistingOptions As WIDGET, psBackdrop As WIDGET, optionsForm As WIDGET, isInBlindMode As Boolean
    Dim pNameEditBox As WIDGET, hostButtonHeight As Long, showHostButton As Boolean
    Set psExistingOptions = widgGetFromID(psWScreen, MULTIOP_OPTIONS)
    If Not psExistingOptions Is Nothing Then
        ' options is already up - refresh data
        updateGameOptions
        Exit Sub
    End If

    widgDelete psWScreen, MULTIOP_OPTIONS                ' clear options list
    widgDelete psWScreen, FRONTEND_SIDETEXT3             ' del text..

    Set psBackdrop = widgGetFromID(psWScreen, FRONTEND_BACKDROP)

    ' draw options box.
    Set optionsForm = IntFormAnimated_new(False)
    psBackdrop.attach optionsForm
    optionsForm.id = MULTIOP_OPTIONS
    optionsForm.setCalcLayout WzFunc_new(FN_optionsForm_calcLayout)

    addSideText FRONTEND_SIDETEXT3, MULTIOP_OPTIONSX - 6, MULTIOP_OPTIONSY, "OPTIONS"

    isInBlindMode = (game.blindMode <> BLIND_MODE_NONE)
    Set pNameEditBox = addMultiEditBox(MULTIOP_OPTIONS, MULTIOP_PNAME, MCOL0, MROW1, "", sPlayer, IMAGE_EDIT_PLAYER, IMAGE_EDIT_PLAYER_HI, MULTIOP_PNAME_ICON, isInBlindMode)
    If Not pNameEditBox Is Nothing Then
        updatePlayerNameBox
    End If

    ' Create the Start Hosting button
    Set startHostingButton = WzStartHostingButton_make(Me)
    W_BUTTON_addOnClickHandler startHostingButton, WzFunc_new(FN_startHosting_onClick, ObjPtr(Me))
    optionsForm.attach startHostingButton
    hostButtonHeight = startHostingButton.idealHeight()
    startHostingButton.setGeometry 1, optionsForm.height() - hostButtonHeight - 1, optionsForm.width() - 2, hostButtonHeight

    showHostButton = (ingame.side = InGameSide_HOST_OR_SINGLEPLAYER) And (Not challengeActive) And Not (NetPlay_isHost)
    If showHostButton Then
        startHostingButton.show
    Else
        startHostingButton.hide
    End If

    ' Add any relevant factory disabled icons / flags
    updateStructureDisabledFlags

    ' Add game options list
    Set multiLobbyOptionsForm = makeWzMultiLobbyOptionsForm(challengeActive, Me)
    optionsForm.attach multiLobbyOptionsForm
    multiLobbyOptionsForm.setCalcLayout WzFunc_new(FN_multiLobbyOptionsForm_calcLayout, ObjPtr(Me))

    ' cancel
    addMultiBut_screen psWScreen, MULTIOP_OPTIONS, CON_CANCEL, _
                MULTIOP_CANCELX, MULTIOP_CANCELY, _
                iV_GetImageWidth(FrontImages, IMAGE_RETURN), _
                iV_GetImageHeight(FrontImages, IMAGE_RETURN), _
                "Return To Previous Screen", IMAGE_RETURN, IMAGE_RETURN_HI, IMAGE_RETURN_HI
End Sub

' (VB) the multiLobbyOptionsForm's calcLayout lambda reads these private members
Public Function getStartHostingButton() As WIDGET
    Set getStartHostingButton = startHostingButton
End Function

Public Function playerRowY0(ByVal row As Long) As Long
    Dim pBlindWaitingRoom As WIDGET
    If isBlindSimpleLobby(game.blindMode) And Not NetPlay_isHost Then
        ' Get Y position of virtual player row in WzPlayerBlindWaitingRoom widget
        Set pBlindWaitingRoom = widgFormGetFromID(psWScreen.psForm, MULTIOP_BLIND_WAITING_ROOM)
        If pBlindWaitingRoom Is Nothing Then
            playerRowY0 = -1
            Exit Function
        End If
        playerRowY0 = getWzPlayerBlindWaitingRoomPlayerRowY0(pBlindWaitingRoom) + pBlindWaitingRoom.y()
        Exit Function
    End If
    If Not ULT(row, playerRows_size) Then
        playerRowY0 = -1
        Exit Function
    End If
    playerRowY0 = playerRows(row).y()
End Function

Public Sub closeAllChoosers()
    closeColourChooser
    closeTeamChooser

    ' AiChooser, DifficultyChooser, and PositionChooser currently use the same form id, so to avoid a double-delete-later, do it once explicitly here
    widgDelete psInlineChooserOverlayScreen, MULTIOP_AI_FORM
    widgDelete psInlineChooserOverlayScreen, FRONTEND_SIDETEXT2
    aiChooserUp = -1
    difficultyChooserUp = -1
    has_playerSlotSwapChooserUp = False
    widgRemoveOverlayScreen psInlineChooserOverlayScreen
End Sub

' Initializes a chooser, preparing to add it on a single line on player list. This involves removing
' any widgets that may have been positioned on that player's row.
Private Sub initInlineChooser(ByVal player As Long)
    ' remove any choosers already up
    closeAllChoosers

    psInlineChooserOverlayScreen.screenSizeDidChange screenWidth, screenHeight, screenWidth, screenHeight ' trigger a screenSizeDidChange so it can relayout (if screen size changed since last time it was registered as an overlay screen)
    widgRegisterOverlayScreen psInlineChooserOverlayScreen, 1
End Sub

' Initializes a chooser, preparing to replace the "right side". This involves removing
' the player list.
Private Function initRightSideChooser(ByVal sideText As String) As WIDGET
    Dim psParent As WIDGET, chooserParent As WIDGET, aiForm As WIDGET, psSideTextLabel As WIDGET
    ' remove any choosers already up
    closeAllChoosers

    ' delete everything on that player's row,
    widgDelete psWScreen, MULTIOP_PLAYERS
    widgDelete psWScreen, FRONTEND_SIDETEXT2

    widgRegisterOverlayScreen psInlineChooserOverlayScreen, 1

    Set psParent = widgGetFromID(psWScreen, FRONTEND_BACKDROP)
    If psParent Is Nothing Then
        Exit Function
    End If

    Set chooserParent = widgGetFromID(psInlineChooserOverlayScreen, MULTIOP_INLINE_OVERLAY_ROOT_FRM)

    Set aiForm = IntFormAnimated_new(False)
    chooserParent.attach aiForm
    aiForm.id = MULTIOP_AI_FORM
    ' (the std::weak_ptr psWeakParent: a strong capture, the form does not own its parent's screen)
    aiForm.setCalcLayout WzFunc_new(FN_rightSideChooser_calcLayout, psParent, 0)

    Set psSideTextLabel = addSideText_screen(psInlineChooserOverlayScreen, MULTIOP_INLINE_OVERLAY_ROOT_FRM, FRONTEND_SIDETEXT2, MULTIOP_PLAYERSX - 3, MULTIOP_PLAYERSY, sideText)
    If Not psSideTextLabel Is Nothing Then
        psSideTextLabel.setCalcLayout WzFunc_new(FN_rightSideChooser_calcLayout, psParent, -3)
    End If

    Set initRightSideChooser = aiForm
End Function

Public Sub openDifficultyChooser(ByVal player As Long)
    Dim aiForm As WIDGET, difficultyIdx As Long, pDifficultyRow As WIDGET, tipStr As String, difficultyTip As String
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:1467", "openDifficultyChooser", "Host only routine detected for client!": Exit Sub

    Set aiForm = initRightSideChooser("DIFFICULTY")
    If aiForm Is Nothing Then
        debug_ 1472, LOG_ERROR, "openDifficultyChooser", "Failed to initialize right-side chooser?"
        Exit Sub
    End If

    For difficultyIdx = AIDifficulty_EASY To AIDifficulty_INSANE
        Set pDifficultyRow = W_BUTTON_new()
        pDifficultyRow.id = 0 'MULTIOP_DIFFICULTY_CHOOSE_START + difficultyIdx;
        pDifficultyRow.setGeometry 7, (MULTIOP_PLAYERHEIGHT + 5) * (difficultyIdx - 1) + 4, MULTIOP_PLAYERWIDTH + 1, MULTIOP_PLAYERHEIGHT
        tipStr = ""
        Select Case difficultyIdx
        Case AIDifficulty_EASY: tipStr = "Starts disadvantaged"
        Case AIDifficulty_MEDIUM: tipStr = "Plays nice"
        Case AIDifficulty_HARD: tipStr = "No holds barred"
        Case AIDifficulty_INSANE: tipStr = "Starts with advantages"
        End Select
        If ULT(NetPlay_players_ai(player), getAIData_size()) Then
            difficultyTip = getAIData_difficultyTip(NetPlay_players_ai(player), difficultyIdx)
            If difficultyTip <> "" Then
                tipStr = tipStr & vbLf
                tipStr = tipStr & difficultyTip
            End If
        Else
            ASSERT_FAILURE "false", "src/multiint.cpp:1515", "openDifficultyChooser", "Invalid AI (index: %i, num AIs: %zu)", NetPlay_players_ai(player), getAIData_size()
        End If
        pDifficultyRow.setTip tipStr
        Set pDifficultyRow.displayFunction = WzFunc_new(FN_displayDifficulty)
        pDifficultyRow.UserData = difficultyIdx
        ' (new DisplayDifficultyCache(), deleted by setOnDelete: its WzText, which goes with the widget)
        Set pDifficultyRow.pUserData = WzText_new("", font_regular)
        aiForm.attach pDifficultyRow
        W_BUTTON_addOnClickHandler pDifficultyRow, WzFunc_new(FN_difficultyRow_onClick, ObjPtr(Me), difficultyIdx, player)
    Next

    difficultyChooserUp = player
End Sub

Public Sub openAiChooser(ByVal player As Long)
    Dim aiForm As WIDGET, spectatorSlotSupported As Boolean, openSlotSupported As Boolean, aiSlotSupported As Boolean
    Dim sButInit As W_BUTINIT, openCloseOnClickHandler As WzFunc, psButton As WIDGET, psCloseButton As WIDGET, psSpectatorButton As WIDGET
    Dim pAIScrollableList As WIDGET, aiListStartXPos As Long, aiListStartYPos As Long, aiListEntryHeight As Long
    Dim aiListEntryWidth As Long, aiListHeight As Long, emptyInit As W_BUTINIT, aiIdx As Long, pAIRow As WIDGET
    Const topMostYPos As Long = 3
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:1549", "openAiChooser", "Host only routine detected for client!": Exit Sub

    Set aiForm = initRightSideChooser("CHOOSE AI")
    If aiForm Is Nothing Then
        debug_ 1554, LOG_ERROR, "openAiChooser", "Failed to initialize right-side chooser?"
        Exit Sub
    End If

    spectatorSlotSupported = spectatorSlotsSupported() And isSpectatorOnlySlot(player)
    openSlotSupported = (NetPlay_bComms) And Not isSpectatorOnlySlot(player)
    aiSlotSupported = Not isSpectatorOnlySlot(player)

    ' (W_BUTINIT sButInit: a zeroed W_BUTINIT is its default, FontID = font_regular)
    sButInit.formID = MULTIOP_AI_FORM
    sButInit.x = 7
    sButInit.width = MULTIOP_PLAYERWIDTH + 1
    sButInit.height = MULTIOP_PLAYERHEIGHT
    Set sButInit.pDisplay = WzFunc_new(FN_displayAi)
    Set sButInit.initPUserDataFunc = WzFunc_new(FN_initPUserData_DisplayAICache)
    ' (sButInit.onDelete deletes the DisplayAICache: its WzText goes with the widget)

    Set openCloseOnClickHandler = WzFunc_new(FN_aiOpenClose_onClick, ObjPtr(Me), player)

    ' Open button
    If openSlotSupported Then
        sButInit.id = MULTIOP_AI_OPEN
        sButInit.pTip = "Allow human players to join in this slot"
        sButInit.UserData = AI_OPEN
        sButInit.y = topMostYPos    'Top most position
        Set psButton = widgAddButton(psInlineChooserOverlayScreen, sButInit)
        If Not psButton Is Nothing Then
            W_BUTTON_addOnClickHandler psButton, openCloseOnClickHandler
        End If
    End If

    ' Closed button
    sButInit.pTip = "Leave this slot unused"
    sButInit.id = MULTIOP_AI_CLOSED
    sButInit.UserData = AI_CLOSED
    If NetPlay_bComms Then
        sButInit.y = sButInit.y + sButInit.height
    Else
        sButInit.y = topMostYPos 'since we don't have the lone mpbutton, we can start at position 0
    End If
    Set psCloseButton = widgAddButton(psInlineChooserOverlayScreen, sButInit)
    If Not psCloseButton Is Nothing Then
        W_BUTTON_addOnClickHandler psCloseButton, openCloseOnClickHandler
    End If

    ' Spectator button
    If spectatorSlotSupported Then
        sButInit.pTip = "Allow spectators to join in this slot"
        sButInit.id = MULTIOP_AI_SPECTATOR
        sButInit.UserData = AI_OPEN
        sButInit.y = sButInit.y + sButInit.height
        Set psSpectatorButton = widgAddButton(psInlineChooserOverlayScreen, sButInit)
        If Not psSpectatorButton Is Nothing Then
            W_BUTTON_addOnClickHandler psSpectatorButton, openCloseOnClickHandler
        End If
    End If

    If aiSlotSupported Then
        Set pAIScrollableList = ScrollableListWidget_make()
        aiForm.attach pAIScrollableList
        ScrollableListWidget_setBackgroundColor pAIScrollableList, psPalette(WZCOL_TRANSPARENT_BOX)
        aiListStartXPos = sButInit.x
        aiListStartYPos = (sButInit.height + sButInit.y) + 10
        aiListEntryHeight = sButInit.height
        aiListEntryWidth = sButInit.width
        If NetPlay_bComms Then aiListHeight = aiListEntryHeight * 7 Else aiListHeight = aiListEntryHeight * 8
        pAIScrollableList.setCalcLayout WzFunc_new(FN_aiList_calcLayout, aiListStartXPos, aiListStartYPos, aiListEntryWidth, aiListHeight)

        For aiIdx = 0 To getAIData_size() - 1
            Set pAIRow = W_BUTTON_new_init(emptyInit)
            pAIRow.setTip getAIData_tip(aiIdx)
            pAIRow.id = 0 'MULTIOP_AI_START + aiIdx;
            pAIRow.UserData = aiIdx
            pAIRow.setGeometry 0, 0, sButInit.width, sButInit.height
            Set pAIRow.displayFunction = WzFunc_new(FN_displayAi)
            ' (new DisplayAICache(), deleted by setOnDelete: its WzText, which goes with the widget)
            Set pAIRow.pUserData = WzText_new("", font_regular)
            W_BUTTON_addOnClickHandler pAIRow, WzFunc_new(FN_aiRow_onClick, ObjPtr(Me), aiIdx, player)
            ScrollableListWidget_addItem pAIScrollableList, pAIRow
        Next
    End If

    aiChooserUp = player
End Sub

Public Sub openPositionChooser(ByVal player As Long)
    Dim psParent As WIDGET, chooserParent As WIDGET, positionChooser As WIDGET
    ' remove any choosers already up
    closeAllChoosers

    widgRegisterOverlayScreen psInlineChooserOverlayScreen, 1

    Set psParent = widgGetFromID(psWScreen, FRONTEND_BACKDROP)
    If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/multiint.cpp:1981", "openPositionChooser", "Could not find root form": Exit Sub

    Set chooserParent = widgGetFromID(psInlineChooserOverlayScreen, MULTIOP_INLINE_OVERLAY_ROOT_FRM)

    Set positionChooser = WzPositionChooser_make(player, WzPlayerSelectionChangePositionRowFactory, Me)
    chooserParent.attach positionChooser
    positionChooser.id = MULTIOP_AI_FORM
    positionChooser.setCalcLayout WzFunc_new(FN_rightSideChooser_calcLayout, psParent, 0)
End Sub

Public Sub updatePlayers()
    addPlayerBox True
End Sub

Public Sub updateGameOptions()
    Dim showHostButton As Boolean
    updatePlayerNameBox

    If WIDGET_isKindOf(multiLobbyOptionsForm, K_WzMultiLobbyOptionsWidgetBase) Then
        vt_WzMultiLobbyOptionsWidgetBase_refreshData multiLobbyOptionsForm
    End If

    If Not startHostingButton Is Nothing Then
        showHostButton = (ingame.side = InGameSide_HOST_OR_SINGLEPLAYER) And (Not challengeActive) And Not (NetPlay_isHost)
        If showHostButton Then
            startHostingButton.show
        Else
            startHostingButton.hide
        End If
    End If

    If Not multiLobbyOptionsForm Is Nothing Then
        multiLobbyOptionsForm.callCalcLayout
    End If
End Sub

Public Sub openTeamChooser(ByVal player As Long)
    Dim bIsTrueMultiplayerGame As Boolean, i As Long, disallow As Long, currSpectatorInfo As Long
    Dim isSpectator As Boolean, canChangeTeams As Boolean, canKickPlayer As Boolean, canChangeSpectatorStatus As Boolean
    Dim displayMoveToSpectatorsButton As Boolean, displayMoveToPlayersButton As Boolean
    Dim psParentForm As WIDGET, psInlineChooserForm As WIDGET, teamW As Long, teamH As Long, space As Long, spaceDiv As Long
    Dim onClickHandler As WzFunc, imgwidth As Long, imgheight As Long, kickImageX As Long, imgwidth_ban As Long, imgheight_ban As Long
    Dim imgwidth_spec As Long, imgheight_spec As Long, onSpecClickHandler As WzFunc, toolTip As String, psButton As WIDGET
    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < MAX_CONNECTED_PLAYERS", "src/multiint.cpp:2036", "openTeamChooser", "Invalid player: %u", player: Exit Sub

    bIsTrueMultiplayerGame = bMultiPlayer And NetPlay_bComms

    disallow = allPlayersOnSameTeam(player)
    If bIsTrueMultiplayerGame And isHostOrAdmin() Then
        ' allow configuration of all teams in true multiplayer mode (by host/admin), even if they would block the game starting
        ' (i.e. even if all players would be configured to be on the same team)
        disallow = -1
    End If
    currSpectatorInfo = NETGameGetSpectatorInfo()

    isSpectator = NetPlay_players_isSpectator(player)
    canChangeTeams = Not locked.teams And Not isSpectator And alliancesSetTeamsBeforeGame(game.alliance)
    canKickPlayer = player <> selectedPlayer And NetPlay_bComms And NetPlay_isHost And NetPlay_players_allocated(player)
    canChangeSpectatorStatus = Not locked.spectators And NetPlay_bComms And isHumanPlayer(player) And (player <> NetPlay_hostPlayer) And (NetPlay_isHost Or player = selectedPlayer)
    displayMoveToSpectatorsButton = canChangeSpectatorStatus And Not isSpectator
    If displayMoveToSpectatorsButton Then displayMoveToSpectatorsButton = (SpectatorInfo_availableSpectatorSlots(currSpectatorInfo) > 0 Or (NetPlay_isHost And NETcanOpenNewSpectatorSlot()))
    displayMoveToPlayersButton = canChangeSpectatorStatus And isSpectator
    canChangeSpectatorStatus = canChangeSpectatorStatus And (displayMoveToSpectatorsButton Or displayMoveToPlayersButton)

    If Not canChangeTeams And Not canKickPlayer And Not canChangeSpectatorStatus Then
        ' can't change teams / spectator status because player is a spectator and spectators are locked
        Exit Sub
    End If

    debug_ 2064, LOG_NET, "openTeamChooser", "Opened team chooser for %d, current team: %d", player, NetPlay_players_team(player)

    initInlineChooser player

    ' add form.
    'addBlueForm(MULTIOP_PLAYERS, MULTIOP_TEAMCHOOSER_FORM, "", 8, playerBoxHeight(player), MULTIOP_ROW_WIDTH, MULTIOP_TEAMSHEIGHT);
    Set psParentForm = widgGetFromID(psWScreen, MULTIOP_PLAYERS)
    Set psInlineChooserForm = addInlineChooserBlueForm(psInlineChooserOverlayScreen, psParentForm, MULTIOP_TEAMCHOOSER_FORM, "", PLAYERBOX_X0 + 1, MAX_(playerRowY0(player), 0), MULTIOP_ROW_WIDTH, MULTIOP_TEAMSHEIGHT)
    If Not (Not psInlineChooserForm Is Nothing) Then ASSERT_FAILURE "psInlineChooserForm != nullptr", "src/multiint.cpp:2072", "openTeamChooser", "Failed to create form"

    If canChangeTeams Then
        teamW = iV_GetImageWidth(FrontImages, IMAGE_TEAM0)
        teamH = iV_GetImageHeight(FrontImages, IMAGE_TEAM0)
        space = (MULTIOP_ROW_WIDTH - 52) - 4 - teamW * (game.maxPlayers + 1)
        spaceDiv = game.maxPlayers
        If 3 * spaceDiv < space Then space = 3 * spaceDiv

        Set onClickHandler = WzFunc_new(FN_team_onClick, player, ObjPtr(Me))

        ' add the teams, skipping the one we CAN'T be on (if applicable)
        For i = 0 To game.maxPlayers - 1
            If i <> disallow Then
                addMultiButWithClickHandler psInlineChooserForm, MULTIOP_TEAMCHOOSER + i, i * (teamW * spaceDiv + space) \ spaceDiv + 3, 6, teamW, teamH, "Team", IMAGE_TEAM0 + i, IMAGE_TEAM0_HI + i, IMAGE_TEAM0_HI + i, onClickHandler
            End If
            ' may want to add some kind of 'can't do' icon instead of being blank?
        Next
    ElseIf isSpectator Then
        imgwidth = iV_GetImageWidth(FrontImages, IMAGE_SPECTATOR)
        imgheight = iV_GetImageHeight(FrontImages, IMAGE_SPECTATOR)
        addMultiBut_screen psInlineChooserOverlayScreen, MULTIOP_TEAMCHOOSER_FORM, MULTIOP_TEAMCHOOSER_SPECTATOR, 3, 6, imgwidth, imgheight, _
            "Spectator", IMAGE_SPECTATOR, IMAGE_SPECTATOR_HI, IMAGE_SPECTATOR_HI
    End If

    ' add kick & ban buttons
    kickImageX = MULTIOP_ROW_WIDTH
    If canKickPlayer Then
        ' Add "kick" button
        imgwidth = iV_GetImageWidth(FrontImages, IMAGE_NOJOIN)
        imgheight = iV_GetImageHeight(FrontImages, IMAGE_NOJOIN)
        kickImageX = MULTIOP_ROW_WIDTH - imgwidth - 4
        addMultiButWithClickHandler psInlineChooserForm, MULTIOP_TEAMCHOOSER_KICK, kickImageX, 8, imgwidth, imgheight, _
            "Kick player", IMAGE_NOJOIN, IMAGE_NOJOIN, IMAGE_NOJOIN, WzFunc_new(FN_kick_onClick, player, ObjPtr(Me))

        ' Add "ban" button
        imgwidth_ban = iV_GetImageWidth(FrontImages, IMAGE_NOJOIN_FULL)
        imgheight_ban = iV_GetImageHeight(FrontImages, IMAGE_NOJOIN_FULL)
        kickImageX = kickImageX - imgwidth - 4
        addMultiButWithClickHandler psInlineChooserForm, MULTIOP_TEAMCHOOSER_BAN, kickImageX, 8, imgwidth_ban, imgheight_ban, _
            "Ban player", IMAGE_NOJOIN_FULL, IMAGE_NOJOIN_FULL, IMAGE_NOJOIN_FULL, WzFunc_new(FN_ban_onClick, player, ObjPtr(Me))
    End If

    If canChangeSpectatorStatus Then
        ' Add a "make spectator" button (if there are available spectator slots, and this is a player)
        If displayMoveToSpectatorsButton Then
            imgwidth_spec = iV_GetImageWidth(FrontImages, IMAGE_SPECTATOR)
            imgheight_spec = iV_GetImageHeight(FrontImages, IMAGE_SPECTATOR)

            If NetPlay_isHost Then
                Set onSpecClickHandler = WzFunc_new(FN_specHost_onClick, player, ObjPtr(Me))
            Else
                Set onSpecClickHandler = WzFunc_new(FN_specClient_onClick, player, ObjPtr(Me))
            End If

            addMultiButWithClickHandler psInlineChooserForm, MULTIOP_TEAMCHOOSER_SPECTATOR, kickImageX - imgwidth_spec - 4, 6, imgwidth_spec, imgheight_spec, _
                "Move to Spectators", IMAGE_SPECTATOR, IMAGE_SPECTATOR_HI, IMAGE_SPECTATOR_HI, onSpecClickHandler
        ElseIf displayMoveToPlayersButton Then
            imgwidth_spec = iV_GetImageWidth(FrontImages, IMAGE_EDIT_PLAYER)
            imgheight_spec = iV_GetImageHeight(FrontImages, IMAGE_EDIT_PLAYER)

            ' (host: ask the spectator if they are okay with a move from spectator -> player; client: ask the host
            ' for a move from spectator -> player: the same lambda body, SendPlayerSlotTypeRequest(player, false))
            If NetPlay_isHost Then
                Set onSpecClickHandler = WzFunc_new(FN_askToPlay_onClick, player, ObjPtr(Me), 2234)
                toolTip = "Ask Spectator to Play"
            Else
                Set onSpecClickHandler = WzFunc_new(FN_askToPlay_onClick, player, ObjPtr(Me), 2249)
                toolTip = "Ask to Play"
            End If

            Set psButton = addMultiButWithClickHandler(psInlineChooserForm, MULTIOP_TEAMCHOOSER_SPECTATOR, kickImageX - imgwidth_spec - 4, 5, imgwidth_spec, imgheight_spec, _
                toolTip, IMAGE_EDIT_PLAYER, IMAGE_EDIT_PLAYER_HI, IMAGE_EDIT_PLAYER_HI, onSpecClickHandler)
            If Not (Not psButton Is Nothing) Then ASSERT_FAILURE "psButton != nullptr", "src/multiint.cpp:2263", "openTeamChooser", "Null button??"
            WzMultiButton_(psButton).has_drawBlueBorder = True ' HACK: to disable the other hack in WzMultiButton::display() for IMAGE_EDIT_PLAYER (sigh...)
            WzMultiButton_(psButton).drawBlueBorder = False
        End If
    End If

    inlineChooserUp = player
End Sub

Public Sub openColourChooser(ByVal player As Long)
    Dim psParentForm As WIDGET, psInlineChooserForm As WIDGET, flagW As Long, flagH As Long, space As Long, spaceDiv As Long
    Dim i As Long, facW As Long, facH As Long
    If Not ULT(player, MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/multiint.cpp:2273", "openColourChooser", "Invalid player number": Exit Sub
    initInlineChooser player

    ' add form.
    Set psParentForm = widgGetFromID(psWScreen, MULTIOP_PLAYERS)
    Set psInlineChooserForm = addInlineChooserBlueForm(psInlineChooserOverlayScreen, psParentForm, MULTIOP_COLCHOOSER_FORM, "", _
        PLAYERBOX_X0, MAX_(playerRowY0(player), 0), _
        MULTIOP_ROW_WIDTH, MULTIOP_PLAYERHEIGHT * 2 - 8)
    If Not (Not psInlineChooserForm Is Nothing) Then ASSERT_FAILURE "psInlineChooserForm != nullptr", "src/multiint.cpp:2281", "openColourChooser", "Failed to create form"

    ' add the flags
    flagW = iV_GetImageWidth(FrontImages, IMAGE_PLAYERN)
    flagH = iV_GetImageHeight(FrontImages, IMAGE_PLAYERN)
    space = MULTIOP_ROW_WIDTH - 0 - flagW * MAX_PLAYERS_IN_GUI
    spaceDiv = MAX_PLAYERS_IN_GUI
    If 5 * spaceDiv < space Then space = 5 * spaceDiv

    For i = 0 To MAX_PLAYERS_IN_GUI - 1
        addMultiButWithClickHandler psInlineChooserForm, MULTIOP_COLCHOOSER + getPlayerColour(i), _
            i * (flagW * spaceDiv + space) \ spaceDiv + 4, 4, _
            flagW, flagH, _
            getPlayerColourName(i), IMAGE_PLAYERN, IMAGE_PLAYERN_HI, IMAGE_PLAYERN_HI, WzFunc_new(FN_colour_onClick, player, ObjPtr(Me)), getPlayerColour(i)

        If Not safeToUseColour(selectedPlayer, i) Then
            widgSetButtonState psInlineChooserOverlayScreen, MULTIOP_COLCHOOSER + getPlayerColour(i), WBUT_DISABLE
        End If
    Next

    facW = iV_GetImageWidth(FrontImages, IMAGE_FACTION_NORMAL) + 4
    facH = iV_GetImageHeight(FrontImages, IMAGE_PLAYERN) - 4
    space = MULTIOP_ROW_WIDTH - 0 - facW * NUM_FACTIONS
    spaceDiv = NUM_FACTIONS
    If 5 * spaceDiv < space Then space = 5 * spaceDiv

    For i = 0 To NUM_FACTIONS - 1
        addMultiButWithClickHandler psInlineChooserForm, MULTIOP_FACCHOOSER + i, _
            i * (facW * spaceDiv + space) \ spaceDiv + PLAYERBOX_X0, MULTIOP_PLAYERHEIGHT, _
            facW, facH, _
            to_localized_string_FactionID(i), _
            IMAGE_FACTION_NORMAL + i, IMAGE_FACTION_NORMAL_HI + i, IMAGE_FACTION_NORMAL_HI + i, WzFunc_new(FN_faction_onClick, player, ObjPtr(Me))
    Next

    inlineChooserUp = player
End Sub

Public Sub openMapChooser()
    widgDelete psWScreen, MULTIOP_PLAYERS
    widgDelete psWScreen, FRONTEND_SIDETEXT2                 ' del text too,

    debug_ 2365, LOG_WZ, "openMapChooser", "openMapChooser: %s.wrf", MultiCustomMapsPath
    addMultiRequest MultiCustomMapsPath, ".wrf", MULTIOP_MAP, current_numplayers

    If NetPlay_isHost And NetPlay_bComms Then
        sendOptions

        NETsetLobbyOptField game.map, NET_LOBBY_OPT_FIELD_MAPNAME
    End If
End Sub

Public Sub closeColourChooser()
    inlineChooserUp = -1
    widgDelete psInlineChooserOverlayScreen, MULTIOP_COLCHOOSER_FORM
    widgRemoveOverlayScreen psInlineChooserOverlayScreen
End Sub

Public Sub closeTeamChooser()
    inlineChooserUp = -1
    widgDelete psInlineChooserOverlayScreen, MULTIOP_TEAMCHOOSER_FORM
    widgRemoveOverlayScreen psInlineChooserOverlayScreen
End Sub

Public Sub closeAiChooser()
    ' AiChooser and DifficultyChooser currently use the same formID
    ' Just call closeAllChoosers() for now
    closeAllChoosers
End Sub

Public Sub closeDifficultyChooser()
    ' AiChooser and DifficultyChooser currently use the same formID
    ' Just call closeAllChoosers() for now
    closeAllChoosers
End Sub

Public Sub closePositionChooser()
    ' AiChooser / DifficultyChooser / PositionChooser currently use the same formID
    ' Just call closeAllChoosers() for now
    closeAllChoosers
End Sub

Public Sub openPlayerSlotSwapChooser(ByVal playerIndex As Long)
    Dim playersTabButtons As WIDGET, psParent As WIDGET, psInlineChooserOverlayRootForm As WIDGET, chooserParent As WIDGET
    Dim psParentPlayersForm As WIDGET, textHeight As Long, swapContextFormMargin As Long, swapContextFormPadding As Long
    Dim swapContextFormHeight As Long, additionalHeightForSwappingRow As Long, aiForm As WIDGET, positionChooser As WIDGET
    Dim swapContextForm As WIDGET, playerRow As WIDGET, playerRowLabel As WIDGET, playerDescription As String
    Dim playerRowSwitchArrow As WIDGET, psSideTextLabel As WIDGET
    ' remove any choosers already up
    closeAllChoosers

    ' make sure "players" view is visible
    Set playersTabButtons = widgFormGetFromID(psWScreen.psForm, MULTIOP_PLAYERS_TABS)
    If WIDGET_isKindOf(playersTabButtons, K_WzPlayerBoxTabs) Then
        WzPlayerBoxTabs_setDisplayView playersTabButtons, PlayerDisplayView_Players
    End If
    updatePlayers

    widgRegisterOverlayScreen psInlineChooserOverlayScreen, 1

    Set psParent = widgGetFromID(psWScreen, FRONTEND_BACKDROP)
    If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/multiint.cpp:3490", "openPlayerSlotSwapChooser", "Could not find root form": Exit Sub

    Set psInlineChooserOverlayRootForm = widgGetFromID(psInlineChooserOverlayScreen, MULTIOP_INLINE_OVERLAY_ROOT_FRM)
    If Not (Not psInlineChooserOverlayRootForm Is Nothing) Then ASSERT_FAILURE "psInlineChooserOverlayRootForm != nullptr", "src/multiint.cpp:3494", "openPlayerSlotSwapChooser", "Could not find overlay root form": Exit Sub
    If WIDGET_isKindOf(psInlineChooserOverlayRootForm, K_W_FULLSCREENOVERLAY_CLICKFORM) Then Set chooserParent = psInlineChooserOverlayRootForm
    Set psParentPlayersForm = widgGetFromID(psWScreen, MULTIOP_PLAYERS)
    If Not (Not psParentPlayersForm Is Nothing) Then ASSERT_FAILURE "psParentPlayersForm != nullptr", "src/multiint.cpp:3497", "openPlayerSlotSwapChooser", "Could not find players form?": Exit Sub
    W_FULLSCREENOVERLAY_CLICKFORM_setCutoutWidget chooserParent, psParentPlayersForm ' should be cleared on close

    textHeight = iV_GetTextLineSize(font_regular)
    swapContextFormMargin = 1
    swapContextFormPadding = 4
    swapContextFormHeight = swapContextFormPadding + textHeight + 4 + MULTIOP_PLAYERHEIGHT + (swapContextFormPadding * 2)
    additionalHeightForSwappingRow = swapContextFormMargin + swapContextFormHeight

    Set aiForm = WIDGET_new()
    chooserParent.attach aiForm
    aiForm.id = MULTIOP_AI_FORM
    ' (psWeakParent: a strong capture, the form does not own its parent's screen)
    aiForm.setCalcLayout WzFunc_new(FN_playerSlotSwapChooser_calcLayout, 0, psParent, additionalHeightForSwappingRow)   ' SWAPLAYOUT_AIFORM
    aiForm.setTransparentToClicks True

    Set positionChooser = WzPositionChooser_make(playerIndex, WzPlayerIndexSwapPositionRowRactory, Me)
    aiForm.attach positionChooser
    positionChooser.setCalcLayout WzFunc_new(FN_playerSlotSwapChooser_calcLayout, 1)                                     ' SWAPLAYOUT_POSITIONCHOOSER

    ' Create parent form for row being swapped
    Set swapContextForm = IntFormAnimated_new(False)
    positionChooser.attach swapContextForm
    swapContextForm.setCalcLayout WzFunc_new(FN_playerSlotSwapChooser_calcLayout, 2, WzPositionChooser_(positionChooser).weakSelf, swapContextFormHeight, swapContextFormMargin) ' SWAPLAYOUT_CONTEXTFORM

    ' Now create a dummy row for the row being swapped, and display beneath
    Set playerRow = WzPlayerRow_make(playerIndex, Me)
    playerRow.setTransparentToMouse True ' ensure clicks on this display-only row have no effect
    swapContextForm.attach playerRow
    playerRow.setCalcLayout WzFunc_new(FN_playerSlotSwapChooser_calcLayout, 3, swapContextFormPadding, textHeight)    ' SWAPLAYOUT_PLAYERROW
    Set playerRowLabel = W_LABEL_new()
    W_LABEL_setFont playerRowLabel, font_regular, psPalette(WZCOL_TEXT_BRIGHT)
    If NetPlay_players_isSpectator(playerIndex) Then playerDescription = "For Spectator:" Else playerDescription = "For Player:"
    playerRowLabel.setString playerDescription
    swapContextForm.attach playerRowLabel
    playerRowLabel.setCalcLayout WzFunc_new(FN_playerSlotSwapChooser_calcLayout, 4, swapContextFormPadding, textHeight) ' SWAPLAYOUT_ROWLABEL
    Set playerRowSwitchArrow = W_LABEL_new()
    W_LABEL_setFont playerRowSwitchArrow, font_regular, psPalette(WZCOL_TEXT_MEDIUM)
    playerRowSwitchArrow.setString ChrW$(&H21C5) ' (up down arrows)
    playerRowSwitchArrow.setGeometry 0, 0, playerRowSwitchArrow.idealWidth(), playerRowSwitchArrow.idealHeight()
    swapContextForm.attach playerRowSwitchArrow
    playerRowSwitchArrow.setCalcLayout WzFunc_new(FN_playerSlotSwapChooser_calcLayout, 5, swapContextFormPadding, textHeight) ' SWAPLAYOUT_SWITCHARROW

    ' Add side text
    Set psSideTextLabel = addSideText_screen(psInlineChooserOverlayScreen, MULTIOP_INLINE_OVERLAY_ROOT_FRM, FRONTEND_SIDETEXT2, MULTIOP_PLAYERSX - 3, MULTIOP_PLAYERSY - 15, "Choose Player Slot")
    If Not psSideTextLabel Is Nothing Then
        psSideTextLabel.setCalcLayout WzFunc_new(FN_playerSlotSwapChooser_calcLayout, 6, psParent)                         ' SWAPLAYOUT_SIDETEXT
    End If

    has_playerSlotSwapChooserUp = True
    playerSlotSwapChooserUp = playerIndex
End Sub

Public Sub closePlayerSlotSwapChooser()
    Dim psInlineChooserOverlayRootForm As WIDGET
    Set psInlineChooserOverlayRootForm = widgGetFromID(psInlineChooserOverlayScreen, MULTIOP_INLINE_OVERLAY_ROOT_FRM)
    If WIDGET_isKindOf(psInlineChooserOverlayRootForm, K_W_FULLSCREENOVERLAY_CLICKFORM) Then
        W_FULLSCREENOVERLAY_CLICKFORM_setCutoutWidget psInlineChooserOverlayRootForm, Nothing
    End If

    ' AiChooser / DifficultyChooser / PositionChooser / PlayerSlotSwapChooser currently use the same formID
    ' Just call closeAllChoosers() for now
    closeAllChoosers

    has_playerSlotSwapChooserUp = False
End Sub

' optional<uint32_t> playerSlotSwapChooserOpenForPlayer() const: has_value() returned, the value in value
Public Function playerSlotSwapChooserOpenForPlayer(value As Long) As Boolean
    playerSlotSwapChooserOpenForPlayer = has_playerSlotSwapChooserUp
    value = playerSlotSwapChooserUp
End Function

' ////////////////////////////////////////////////////////////////////////////
' box for players.

Public Sub addPlayerBox(ByVal players As Boolean)
    Dim widgetParent As WIDGET, playersForm As WIDGET, pPlayersLabel As WIDGET, blindWaitingRoom As WIDGET
    Dim isMultiplayMode As Boolean, playersTabButtons As WIDGET, playerDisplayView As WzPlayerBoxTabs_PlayerDisplayView
    Dim numSlotsDisplayed As Long, i As Long, playerRow As WIDGET, playerRowPosition As Long, spectatorAddButton As WIDGET
    Dim imgheight As Long, sAddSpecSlotInit As W_BUTINIT, pButton As WIDGET
    ' if background isn't there, then return since were not ready to draw the box yet!
    If widgGetFromID(psWScreen, FRONTEND_BACKDROP) Is Nothing Then
        Exit Sub
    End If

    If aiChooserUp >= 0 Then
        Exit Sub
    ElseIf difficultyChooserUp >= 0 Then
        Exit Sub
    End If

    ' draw player window
    Set widgetParent = widgGetFromID(psWScreen, FRONTEND_BACKDROP)
    Set playersForm = widgFormGetFromID(psWScreen.psForm, MULTIOP_PLAYERS)
    If Not WIDGET_isKindOf(playersForm, K_IntFormAnimated) Then Set playersForm = Nothing
    If playersForm Is Nothing Then
        widgDelete psWScreen, MULTIOP_PLAYERS
        Set playersForm = IntFormAnimated_new(False)
        widgetParent.attach playersForm
        playersForm.id = MULTIOP_PLAYERS
        playersForm.setCalcLayout WzFunc_new(FN_playersForm_calcLayout)
    End If

    widgDelete psWScreen, FRONTEND_SIDETEXT2  ' del text
    Set pPlayersLabel = addSideText(FRONTEND_SIDETEXT2, MULTIOP_PLAYERSX - 3, MULTIOP_PLAYERSY, "PLAYERS")
    pPlayersLabel.hide ' hide for now

    If Not players Then
        Exit Sub
    End If

    If isBlindSimpleLobby(game.blindMode) And Not NetPlay_isHost Then
        ' Display custom "waiting room" that completely blinds the current player

        Set blindWaitingRoom = makeWzPlayerBlindWaitingRoom(Me)
        blindWaitingRoom.id = MULTIOP_BLIND_WAITING_ROOM
        playersForm.attach blindWaitingRoom
        blindWaitingRoom.setCalcLayout WzFunc_new(FN_blindWaitingRoom_calcLayout)
    Else
        ' Normal case - display players / spectators lists

        ' Add "players" / "spectators" tab buttons (if in multiplay mode)
        isMultiplayMode = NetPlay_bComms And (NetPlay_isHost Or ingame.side = InGameSide_MULTIPLAYER_CLIENT)
        Set playersTabButtons = widgFormGetFromID(psWScreen.psForm, MULTIOP_PLAYERS_TABS)
        If Not WIDGET_isKindOf(playersTabButtons, K_WzPlayerBoxTabs) Then Set playersTabButtons = Nothing
        If isMultiplayMode And playersTabButtons Is Nothing Then
            Set playersTabButtons = WzPlayerBoxTabs_make(NetPlay_isHost, Me)
            playersTabButtons.id = MULTIOP_PLAYERS_TABS
            playersForm.attach playersTabButtons
            playersTabButtons.setCalcLayout WzFunc_new(FN_playersTabs_calcLayout)
        End If
        If Not playersTabButtons Is Nothing Then playerDisplayView = WzPlayerBoxTabs_getDisplayView(playersTabButtons) Else playerDisplayView = PlayerDisplayView_Players

        If Not (MAX_CONNECTED_PLAYERS <= NetPlay_players_size()) Then ASSERT_FAILURE "static_cast<size_t>(((11 + 2) + 10)) <= NetPlay.players.size()", "src/multiint.cpp:3679", "addPlayerBox", "Insufficient array size: %zu versus %zu", NetPlay_players_size(), MAX_CONNECTED_PLAYERS
        numSlotsDisplayed = 0
        If playerRows_size = 0 Then
            For i = 0 To MAX_CONNECTED_PLAYERS - 1
                If Not ingame.localOptionsReceived Then
                    Exit For
                End If

                ' Create player row
                Set playerRow = WzPlayerRow_make(i, Me)
                playersForm.attach playerRow
                ReDim Preserve playerRows(0 To playerRows_size)
                Set playerRows(playerRows_size) = playerRow
                playerRows_size = playerRows_size + 1
            Next
        End If

        For i = 0 To MAX_CONNECTED_PLAYERS - 1
            If Not ingame.localOptionsReceived Then
                Exit For
            End If
            If i >= playerRows_size Then
                If Not (i < playerRows_size) Then ASSERT_FAILURE "i < playerRows.size()", "src/multiint.cpp:3705", "addPlayerBox", "Row widget does not currently exist"
                GoTo continue_
            End If
            Set playerRow = playerRows(i)
            If Not WIDGET_isKindOf(playerRow, K_WzPlayerRow) Then Set playerRow = Nothing

            Select Case playerDisplayView
            Case PlayerDisplayView_Players
                If isSpectatorOnlySlot(i) Then
                    ' should not be visible
                    playerRow.hide
                    GoTo continue_ ' skip
                End If
            Case PlayerDisplayView_Spectators
                If Not isSpectatorOnlySlot(i) Then
                    ' should not be visible
                    playerRow.hide
                    GoTo continue_ ' skip
                End If
            End Select

            If isSpectatorOnlySlot(i) Then
                If (i >= NetPlay_players_size()) Or Not (NetPlay_players_isSpectator(i) And NetPlay_players_ai(i) = AI_OPEN) Then
                    ' the only slots displayable beyond game.maxPlayers are spectator slots
                    playerRow.hide
                    GoTo continue_
                End If
            End If

            If playerDisplayView = PlayerDisplayView_Players Then playerRowPosition = NetPlay_players_position(i) Else playerRowPosition = numSlotsDisplayed
            playerRow.setGeometry PLAYERBOX_X0, playerBoxHeight(playerRowPosition), MULTIOP_PLAYERWIDTH, MULTIOP_PLAYERHEIGHT
            playerRow.show
            WzPlayerRow_updateState playerRow
            ' make sure it's attached to the current player form
            If playerRow.parent() Is Nothing Then
                playersForm.attach playerRow
            End If

            numSlotsDisplayed = numSlotsDisplayed + 1
continue_:
        Next

        Set spectatorAddButton = widgFormGetFromID(psWScreen.psForm, MULTIOP_ADD_SPECTATOR_SLOTS)
        If Not WIDGET_isKindOf(spectatorAddButton, K_W_BUTTON) Then Set spectatorAddButton = Nothing
        If spectatorAddButton Is Nothing Then
            ' Create a "+ Spectator Slots" button
            imgheight = iV_GetImageHeight(FrontImages, IMAGE_SPECTATOR) + 2
            sAddSpecSlotInit.formID = MULTIOP_PLAYERS
            sAddSpecSlotInit.id = MULTIOP_ADD_SPECTATOR_SLOTS
            sAddSpecSlotInit.x = 7
            sAddSpecSlotInit.y = MULTIOP_PLAYERSH - (imgheight + 4) - 4
            sAddSpecSlotInit.width = MULTIOP_TEAMSWIDTH
            sAddSpecSlotInit.height = imgheight
            sAddSpecSlotInit.pTip = "Add spectator slot"
            Set sAddSpecSlotInit.pDisplay = WzFunc_new(FN_displaySpectatorAddButton)
            sAddSpecSlotInit.UserData = 0
            Set pButton = widgAddButton(psWScreen, sAddSpecSlotInit)
            If Not pButton Is Nothing Then
                W_BUTTON_addOnClickHandler pButton, WzFunc_new(FN_addSpectatorSlots_onClick, ObjPtr(Me))
            End If
            Set spectatorAddButton = pButton
        End If

        If Not (Not spectatorAddButton Is Nothing) Then ASSERT_FAILURE "spectatorAddButton != nullptr", "src/multiint.cpp:3788", "addPlayerBox", "Unable to create or find button": Exit Sub
        If (playerDisplayView = PlayerDisplayView_Spectators) And (numSlotsDisplayed < (MAX_PLAYERS_IN_GUI)) And spectatorSlotsSupported() And NetPlay_isHost Then
            ' spectator add button should be visible
            spectatorAddButton.show
        Else
            spectatorAddButton.hide
        End If
    End If
End Sub

' Starts hosting. Used as an action for "Host Game" button. Another place where this is used is for
' activating the host immediately when entering the multiplayer menu with a challenge active.
Public Function startHost() As Boolean
    Dim i As Long, bIsAutoHostOrAutoGame As Boolean
    debug_ 5377, LOG_NET, "startHost", "MULTIOP_HOST enabled"
    resetLobbyChangeVoteData
    resetDataHash
    resetReadyStatus False
    removeWildcards sPlayer
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        ingame.PingTimes(i) = 0
        ingame.VerifiedIdentity(i) = False
        ingame.LagCounter(i) = 0
        ingame.DesyncCounter(i) = 0
        ingame.has_lastSentPlayerDataCheck2(i) = False
        ingame.muteChat(i) = False
    Next
    resetAllMultiOptionPrefValues
    multiSyncResetAllChallenges

    If game.blindMode <> BLIND_MODE_NONE Then
        generateBlindIdentity
    End If

    bIsAutoHostOrAutoGame = getHostLaunch() = HostLaunch_Skirmish Or getHostLaunch() = HostLaunch_Autohost
    If Not hostCampaign(game.name_, sPlayer, spectatorHost, bIsAutoHostOrAutoGame, defaultOpenSpectatorSlots) Then
        displayRoomSystemMessage "Sorry! Failed to host the game."
        wz_command_interface_output "WZEVENT: hostingFailed" & vbLf
        If bIsAutoHostOrAutoGame And headlessGameMode() Then
            wzQuit 1
        End If
        startHost = False
        Exit Function
    End If

    If NET_getLobbyDisabled() Then
        displayLobbyDisabledNotification
    End If

    updateInActualHostedLobby True

    widgDelete psWScreen, MULTIOP_REFRESH
    widgDelete psWScreen, MULTIOP_FILTER_TOGGLE

    ingame.localOptionsReceived = True

    addGameOptions
    addChatBox

    addPlayerBox True

    If game.blindMode <> BLIND_MODE_NONE Then
        printBlindModeHelpMessagesToConsole
    End If

    startHost = True
End Function

'
' Process click events on the multiplayer/skirmish options screen
' 'id' is id of the button that was pressed
'
Public Sub processMultiopWidgets(ByVal id As Long)
    Dim playerStats As PLAYERSTATS_, sPlayer_new As String
    ' these work all the time.
    Select Case id

    Case MULTIOP_PNAME
        sPlayer_new = sstrcpy_cut(widgGetString(psWScreen, MULTIOP_PNAME), 128)

        ' chop to 15 chars..
        Do While Len(sPlayer_new) > 15    ' clip name.
            sPlayer_new = Left$(sPlayer_new, Len(sPlayer_new) - 1)
        Loop
        removeWildcards sPlayer_new

        playerStats = getMultiStats(selectedPlayer)        ' retrieve current identity + stats to pass-in to loadMultiStats on name change
        If loadMultiStats(sPlayer_new, playerStats) Then
            sPlayer = sPlayer_new
            ' update string.
            widgSetString psWScreen, MULTIOP_PNAME, sPlayer
            printConsoleNameChange NetPlay_players_name(selectedPlayer), sPlayer

            NETchangePlayerName selectedPlayer, sPlayer        ' update if joined.
            setMultiStats selectedPlayer, playerStats, False
            setMultiStats selectedPlayer, playerStats, True
            netPlayersUpdated = True

            If NetPlay_isHost And NetPlay_bComms Then
                sendOptions
                NETsetLobbyOptField NetPlay_players_name(selectedPlayer), NET_LOBBY_OPT_FIELD_HOSTNAME
            End If
        Else
            debug_ 5482, LOG_ERROR, "processMultiopWidgets", "Failed to load or create player profile: %s", sPlayer_new
        End If

    Case MULTIOP_PNAME_ICON
        widgDelete psWScreen, MULTIOP_PLAYERS
        widgDelete psWScreen, FRONTEND_SIDETEXT2                  ' del text too,

        addMultiRequest MultiPlayersPath, ".sta2", MULTIOP_PNAME, 0

    Case MULTIOP_CHATEDIT
        ' now handled in setOnReturnHandler

    Case CON_CANCEL

        resetHostLaunch ' Dont load the autohost file on subsequent hosts
        performedFirstStart = False ' Reset everything
        If Not challengeActive Then
            If NetPlay_bComms And ingame.side = InGameSide_MULTIPLAYER_CLIENT And Not NetPlay_isHost Then
                ' remove a potential "allow" vote if we gracefully leave
                sendLobbyChangeVoteData 0
            End If
            NETGameLocked False       ' reset status on a cancel
            stopJoining parent, ERROR_NOERROR
        Else
            NETclose
            ingame.localJoiningInProgress = False
            changeTitleMode SINGLE_
            pie_LoadBackDrop SCREEN_RANDOMBDROP
            addChallenges
        End If
    Case Else
    End Select
End Sub

Public Function getOption_SpectatorHost() As Boolean
    getOption_SpectatorHost = spectatorHost
End Function

Public Sub setOption_SpectatorHost(ByVal value As Boolean)
    spectatorHost = value
    If Not spectatorHost Then
        ' disable blind mode options (which are only available when spectator host)
        game.blindMode = BLIND_MODE_NONE
    End If
End Sub

' (VB) client-only: the redirect join (parseKickRedirectInfo, NET_getCurrentHostTextAddress, the verified join
' identity, startJoinRedirectAttempt) comes with the joining, cluster J; its checks and the failure path are kept
Public Sub handleKickRedirect(ByVal kickerPlayerIdx As Long, ByVal redirectString As String)
    If Not (Not NetPlay_isHost) Then ASSERT_FAILURE "!NetPlay.isHost", "src/multiint.cpp:5917", "handleKickRedirect", "Host shouldn't be processing kick redirect": Exit Sub
    If Not ((GetGameMode() <> GS_NORMAL)) Then ASSERT_FAILURE "(GetGameMode() != GS_NORMAL)", "src/multiint.cpp:5918", "handleKickRedirect", "Game must not be started": Exit Sub

    If kickerPlayerIdx < NetPlay_players_size() Then
        ActivityManager_wasKickedByPlayer kickerPlayerIdx, ERROR_REDIRECT, redirectString
    End If

    debug_ 5931, LOG_ERROR, "handleKickRedirect", "Unable to get current host's address?"
    stopJoining WzMsgBoxTitleUI_new("Disconnected from host:", "Unable to handle host redirect", parent), ERROR_REDIRECT
End Sub

Public Function frontendMultiMessages(ByVal running As Boolean) As MultiMessagesResult
    Dim queue As Long, typ As Long, ignoredMessage As Boolean, done As Boolean, hash As Sha256, r As MessageReader
    Dim minAutoStartPlayerCount As Long, msg As String, player_id As Long, randomSeed As Long, reason As String
    Dim KICK_TYPE As Long, buf As String, kickReasonStr As String, maxLinePos As Long, message As NetworkTextMessage
    Dim displayedMessage As Boolean, isLobbySlashCommand As Boolean, result As SwapPlayerIndexesResult, mutedUntil As Currency
    Dim rm As RoomMessage, now_ As Currency
    ignoredMessage = False

    Do While NETrecvNet(queue, typ)
        If Not shouldProcessMessage(queue, typ) Then
            GoTo continue_
        End If

        ' Copy the message to the global one used by the new NET API
        Select Case typ
        Case NET_FILE_REQUESTED
            recvMapFileRequested queue

        Case NET_FILE_PAYLOAD
            If NetPlay_hostPlayer <> NETQUEUE_index(queue) Then
                HandleBadParam "NET_FILE_PAYLOAD given incorrect params.", 255, NETQUEUE_index(queue)
                ignoredMessage = True
            Else
                done = recvMapFileData(queue)
                If done Then
                    refreshMultiplayerOptionsTitleUIIfActive 'refresh to see the proper tech level in the map name
                End If
                ' spectators should automatically become ready as soon as necessary files are downloaded
                ' and not-ready when files remain to be downloaded
                handleAutoReadyRequest
            End If

        Case NET_FILE_CANCELLED
            If Not NetPlay_isHost Then
                ASSERT_FAILURE "false", "src/multiint.cpp:6009", "frontendMultiMessages", "Host only routine detected for client!"
            Else
                Set r = NETbeginDecode(queue, NET_FILE_CANCELLED)
                NETbin r, hash.bytes, 32
                NETend r

                debug_ 6018, LOG_WARNING, "frontendMultiMessages", "Received file cancel request from player %u, they weren't expecting the file.", NETQUEUE_index(queue)
                ' (the player's wzFiles, erased of the file with that hash: the map downloads, cluster J)
                NetPlay_players_wzFiles_eraseHash NETQUEUE_index(queue), hash
            End If

        Case NET_OPTIONS                    ' incoming options file.
            If NetPlay_hostPlayer <> NETQUEUE_index(queue) Then
                HandleBadParam "NET_OPTIONS should be sent by host", 255, NETQUEUE_index(queue)
                ignoredMessage = True
            ElseIf Not recvOptions(queue) Then
                ' supplied NET_OPTIONS are not valid
                stopJoining WzMsgBoxTitleUI_new("Disconnected from host:", "Host supplied invalid options", parent), ERROR_INVALID
                frontendMultiMessages = MultiMessagesResult_StoppedJoining
                Exit Function
            Else
                updateInActualHostedLobby True
                ingame.localOptionsReceived = True

                handleAutoReadyRequest

                If TypeOf wzTitleUICurrent Is WzMultiplayerOptionsTitleUI Then
                    addGameOptions
                    addChatBox
                End If
            End If

        Case NET_HOST_CONFIG
            If NetPlay_hostPlayer <> NETQUEUE_index(queue) Then
                HandleBadParam "NET_HOST_CONFIG should be sent by host", 255, NETQUEUE_index(queue)
                ignoredMessage = True
            ElseIf Not recvHostConfig(queue) Then
                ' supplied NET_HOST_CONFIG is not valid
                stopJoining WzMsgBoxTitleUI_new("Disconnected from host:", "Host supplied invalid host config", parent), ERROR_INVALID
            ElseIf TypeOf wzTitleUICurrent Is WzMultiplayerOptionsTitleUI Then
                addChatBox ' refresh chat box options
            End If

        Case GAME_ALLIANCE
            recvAlliance queue, False

        Case NET_COLOURREQUEST
            If multiplayIsStartingGame() Then
                ignoredMessage = True
            Else
                recvColourRequest queue
            End If

        Case NET_FACTIONREQUEST
            If multiplayIsStartingGame() Then
                ignoredMessage = True
            Else
                recvFactionRequest queue
            End If

        Case NET_POSITIONREQUEST
            If multiplayIsStartingGame() Then
                ignoredMessage = True
            Else
                recvPositionRequest queue
            End If

        Case NET_TEAMREQUEST
            If multiplayIsStartingGame() Then
                ignoredMessage = True
            Else
                recvTeamRequest queue
            End If

        Case NET_READY_REQUEST
            If multiplayIsStartingGame() Then
                ignoredMessage = True
            Else
                recvReadyRequest queue

                ' If hosting and game not yet started, try to start the game if everyone is ready.
                If NetPlay_isHost And multiplayPlayersReady() Then
                    If Not multiplayLacksEnoughPlayersToAutostart() Then
                        startMultiplayerGame

                        ' Start the game before processing more messages.
                        ' (startMultiplayerGame transitions to new TitleUI)
                        NETpop queue
                        frontendMultiMessages = MultiMessagesResult_StoppedJoining
                        Exit Function
                    Else
                        minAutoStartPlayerCount = getBoundedMinAutostartPlayerCount()
                        If minAutoStartPlayerCount > 0 Then
                            msg = va("Game will not start until there are %d players.", minAutoStartPlayerCount)
                            sendRoomSystemMessage msg
                        End If
                    End If
                End If
            End If

        Case NET_PING                       ' diagnostic ping msg.
            recvPing queue

        Case NET_PLAYER_DROPPED         ' remote player got disconnected
            player_id = MAX_CONNECTED_PLAYERS

            resetReadyStatus False, shouldSkipReadyResetOnPlayerJoinLeaveEvent()

            Set r = NETbeginDecode(queue, NET_PLAYER_DROPPED)
            NETuint32_t r, player_id
            NETend r

            If player_id >= MAX_CONNECTED_PLAYERS Then
                debug_ 6168, LOG_INFO, "frontendMultiMessages", "** player %u has dropped - huh?", player_id
            ElseIf whosResponsible(player_id) <> NETQUEUE_index(queue) And NETQUEUE_index(queue) <> NetPlay_hostPlayer Then
                HandleBadParam "NET_PLAYER_DROPPED given incorrect params.", player_id, NETQUEUE_index(queue)
            Else
                debug_ 6178, LOG_INFO, "frontendMultiMessages", "** player %u has dropped!", player_id

                MultiPlayerLeave player_id      ' get rid of their stuff
                NET_InitPlayer player_id, False ' sets index player's array to false
                NETsetPlayerConnectionStatus CONNECTIONSTATUS_PLAYER_DROPPED, player_id
                If player_id < MAX_PLAYERS Then
                    resetLobbyChangePlayerVote player_id
                End If
                resetMultiOptionPrefValues player_id
                ActivityManager_updateMultiplayGameData
                If player_id = NetPlay_hostPlayer Or player_id = selectedPlayer Then ' if host quits or we quit, abort out
                    stopJoining parent, ERROR_NOERROR
                End If
                updateGameOptions
            End If

        Case NET_PLAYERRESPONDING           ' remote player is now playing.
            Set r = NETbeginDecode(queue, NET_PLAYERRESPONDING)
            ' the player that has just responded
            NETuint32_t r, player_id
            NETend r

            If player_id >= MAX_CONNECTED_PLAYERS Then
                debug_ 6207, LOG_ERROR, "frontendMultiMessages", "Bad NET_PLAYERRESPONDING received, ID is %d", player_id
            ElseIf whosResponsible(player_id) <> NETQUEUE_index(queue) And NETQUEUE_index(queue) <> NetPlay_hostPlayer Then
                HandleBadParam "NET_PLAYERRESPONDING given incorrect params.", player_id, NETQUEUE_index(queue)
            Else
                ingame.JoiningInProgress(player_id) = False
                ingame.DataIntegrity(player_id) = False
            End If

        Case NET_FIREUP                     ' campaign game started.. can fire the whole shebang up...
            cancelOrDismissVoteNotifications ' don't need vote notifications anymore
            cancelOrDismissNotificationsWithTag LOBBY_DISABLED_TAG
            cancelOrDismissNotificationIfTagHasPrefix SLOTTYPE_TAG_PREFIX ' (the lambda: tag.rfind(SLOTTYPE_TAG_PREFIX, 0) == 0)
            If NetPlay_hostPlayer <> NETQUEUE_index(queue) Then
                HandleBadParam "NET_FIREUP given incorrect params.", 255, NETQUEUE_index(queue)
            Else
                debug_ 6232, LOG_INFO, "frontendMultiMessages", "NET_FIREUP was received ..."
                If ingame.localOptionsReceived Then
                    randomSeed = 0
                    Set r = NETbeginDecode(queue, NET_FIREUP)
                    NETuint32_t r, randomSeed
                    NETend r

                    saveMultiOptionPrefValues sPlayer, selectedPlayer ' persist any changes to multioption preferences

                    gameSRand randomSeed  ' Set the seed for the synchronised random number generator, using the seed given by the host.

                    debug_ 6244, LOG_NET, "frontendMultiMessages", "& local Options Received (MP game)"
                    ingame.has_TimeEveryoneIsInGame = False         ' reset time
                    ingame.has_endTime = False
                    resetDataHash
                    decideWRF

                    bMultiPlayer = True
                    bMultiMessages = True
                    changeTitleMode STARTGAME

                    ' Start the game before processing more messages.
                    NETpop queue
                    frontendMultiMessages = MultiMessagesResult_StoppedJoining
                    Exit Function
                End If
                ASSERT_FAILURE "false", "src/multiint.cpp:6258", "frontendMultiMessages", "NET_FIREUP was received, but !ingame.localOptionsReceived."
            End If

        Case NET_KICK                       ' player is forcing someone to leave
            player_id = MAX_CONNECTED_PLAYERS
            KICK_TYPE = ERROR_NOERROR

            Set r = NETbeginDecode(queue, NET_KICK)
            NETuint32_t r, player_id
            NETstring r, reason, MAX_KICK_REASON
            NETenum r, KICK_TYPE
            NETend r

            If player_id >= MAX_CONNECTED_PLAYERS Then
                debug_ 6275, LOG_ERROR, "frontendMultiMessages", "NET_KICK message with invalid player_id: (%u)", player_id
                GoTo endSelect_
            End If

            If player_id < MAX_PLAYERS Then
                resetLobbyChangePlayerVote player_id
            End If

            If player_id = NetPlay_hostPlayer Then
                buf = va("*Player %d (%s : %s) tried to kick %u", NETQUEUE_index(queue), getPlayerName(NETQUEUE_index(queue), True), NetPlay_players_IPtextAddress(NETQUEUE_index(queue)), player_id)
                NETlogEntry buf, SYNC_FLAG, 0
                debug_ 6290, LOG_ERROR, "frontendMultiMessages", "%s", buf
                If NetPlay_isHost Then
                    NETplayerKicked NETQUEUE_index(queue)
                End If
                GoTo endSelect_
            End If

            If selectedPlayer = player_id Then  ' we've been told to leave.
                kickReasonStr = reason
                If KICK_TYPE <> ERROR_REDIRECT Then
                    ' normal kick
                    maxLinePos = nthOccurrenceOfChar(kickReasonStr, vbLf, 10)
                    If maxLinePos <> -1 Then        ' std::string::npos
                        kickReasonStr = Left$(kickReasonStr, maxLinePos)
                    End If
                    stopJoining WzMsgBoxTitleUI_new("You have been kicked: ", kickReasonStr, parent), KICK_TYPE
                    debug_ 6310, LOG_INFO, "frontendMultiMessages", "You have been kicked, because %s ", kickReasonStr
                    displayKickReasonPopup kickReasonStr
                    ActivityManager_wasKickedByPlayer NETQUEUE_index(queue), KICK_TYPE, reason
                Else
                    ' kick_redirect
                    ' (the kickReasonStr is expected to be a formatted string)
                    NETpop queue
                    handleKickRedirect NETQUEUE_index(queue), kickReasonStr
                    frontendMultiMessages = MultiMessagesResult_StoppedJoining
                    Exit Function
                End If
            Else
                NETplayerKicked player_id, KICK_TYPE = ERROR_REDIRECT
            End If

        Case NET_HOST_DROPPED
            Set r = NETbeginDecode(queue, NET_HOST_DROPPED)
            NETend r
            stopJoining WzMsgBoxTitleUI_new("Connection lost:", "No connection to host.", parent), ERROR_HOSTDROPPED
            debug_ 6334, LOG_NET, "frontendMultiMessages", "The host has quit!"

        Case NET_TEXTMSG                    ' Chat message
            If ingame.localOptionsReceived Then
                If NetworkTextMessage_receive(message, queue) Then

                    displayedMessage = False
                    If message.sender < 0 Then
                        displayedMessage = True
                    ElseIf Not isPlayerMuted(message.sender) And Not playerSpamMutedUntil(message.sender, mutedUntil) Then
                        displayedMessage = True
                    End If
                    If displayedMessage Then
                        rm = buildMessage(message.sender, message.text)
                        displayRoomMessage rm
                        audio_PlayTrack FE_AUDIO_MESSAGEEND

                        If message.sender >= 0 Then
                            recordPlayerMessageSent message.sender
                        End If
                    End If

                    isLobbySlashCommand = False
                    If lobby_slashcommands_enabled() Then
                        isLobbySlashCommand = processChatLobbySlashCommands(message)
                        If isLobbySlashCommand And Not displayedMessage Then
                            ' display it anyway, even though user is muted, because it was a processed slash command
                            rm = buildMessage(message.sender, message.text)
                            displayRoomMessage rm
                            audio_PlayTrack FE_AUDIO_MESSAGEEND
                            displayedMessage = True
                        End If
                    End If

                    If Not isLobbySlashCommand Then
                        cmdInterfaceLogChatMsg message, "WZCHATLOB"
                    End If
                End If
            End If

        Case NET_QUICK_CHAT_MSG
            If ingame.localOptionsReceived Then
                recvQuickChat queue
            End If

        Case NET_VOTE
            If NetPlay_isHost And ingame.localOptionsReceived Then
                If recvVote(queue, True) Then
                    refreshMultiplayerOptionsTitleUIIfActive
                End If
            End If

        Case NET_VOTE_REQUEST
            If Not NetPlay_isHost And Not NetPlay_players_isSpectator(selectedPlayer) Then
                recvVoteRequest queue
            End If

        Case NET_PLAYER_SLOTTYPE_REQUEST
            ' Note: Because this may trigger a player index + queue swap,
            ' it handles NETpop(queue) itself and we must immediately continue to the next loop iteration
            ' i.e. Don't use `queue` after recvPlayerSlotTypeRequestAndPop is called
            recvPlayerSlotTypeRequestAndPop Me, queue
            GoTo continue_ ' loop for next queue that has a message

        Case NET_PLAYER_SWAP_INDEX
            result = recvSwapPlayerIndexes(queue, parent)
            If result <> SwapPlayerIndexesResult_SUCCESS Then
                ignoredMessage = True
                If result = SwapPlayerIndexesResult_ERROR_HOST_MADE_US_PLAYER_WITHOUT_PERMISSION Then
                    ' Leave the badly behaved (likely modified) host!
                    sendRoomChatMessage "The host moved me to Players, but I never gave permission for this change. Bye!"
                    debug_ 6419, LOG_INFO, "frontendMultiMessages", "Leaving game because host moved us to Players, but we never gave permission."
                    stopJoining WzMsgBoxTitleUI_new("Disconnected from host:", "The host tried to move us to Players, but we never gave permission.", parent), ERROR_HOSTDROPPED
                    frontendMultiMessages = MultiMessagesResult_StoppedJoining
                    Exit Function
                End If
            End If

        ' Note: NET_PLAYER_SWAP_INDEX_ACK is handled in netplay

        Case Else
            ignoredMessage = True
        End Select
endSelect_:

        If ignoredMessage Then
            debug_ 6436, LOG_ERROR, "frontendMultiMessages", "Didn't handle %s message!", messageTypeToString(typ)
            ignoredMessage = False
        End If

        NETpop queue
continue_:
    Loop

    QueryPerformanceCounter now_
    autoLagKickRoutine now_
    autoLobbyNotReadyKickRoutine now_

    frontendMultiMessages = MultiMessagesResult_Continue
End Function

Private Function WzTitleUI_run() As Long
    Static lastrefresh As Long          ' static UDWORD lastrefresh = 0;
    Dim playerStats As PLAYERSTATS_, isHeadless As Boolean, pChatBox As WIDGET, triggers() As WidgetTrigger, n As Long
    Dim id As Long, mapData As LEVEL_DATASET, isHoverPreview As Boolean, sTemp As String, sPlayer_new As String
    Dim oldGameMap As String, oldGameHash As Sha256, oldMaxPlayers As Long, oldGameIsMapMod As Boolean, oldGameIsRandom As Boolean
    Dim numHumans As Long, humans() As Long, nHumans As Long, playerInc As Long, slotInc As Long, mapName As String

    isHeadless = headlessGameMode()

    If frontendMultiMessages(True) = MultiMessagesResult_StoppedJoining Then
        ' shortcut any further processing - stopped joining, this title UI is about to go away
        WzTitleUI_run = TITLECODE_CONTINUE
        Exit Function
    End If
    If NetPlay_isHost Then
        ' send it for each player that needs it
        sendMap
    End If

    ' update boxes?
    If netPlayersUpdated Or (NetPlay_isHost And mouseDown(MOUSE_LMB) And gameTime - lastrefresh > 500) Then
        netPlayersUpdated = False
        lastrefresh = gameTime
        If Not multiRequestUp And (NetPlay_isHost Or ingame.localJoiningInProgress) And Not isHeadless Then
            If aiChooserUp >= 0 Then
                If NetPlay_players_allocated(aiChooserUp) Then
                    closeAiChooser ' close ai chooser that's open for a slot which a player just joined into
                End If
            End If

            addPlayerBox True               ' update the player box.
            loadMapPreview False
            updateGameOptions
        End If
    End If


    ' update scores and pings if far enough into the game
    If ingame.localOptionsReceived And ingame.localJoiningInProgress Then
        sendScoreCheck
        sendPing
    End If

    ' if we don't have the focus, then autoclick in the chatbox.
    If Not isHeadless And psWScreen.psFocus Is Nothing And Not isMouseOverScreenOverlayChild(mouseX(), mouseY()) And Not mouseDown(MOUSE_LMB) Then
        Set pChatBox = widgGetFromID(psWScreen, MULTIOP_CHATBOX)
        If WIDGET_isKindOf(pChatBox, K_ChatBoxWidget) Then
            If wzSeemsLikeNonTouchPlatform() Then ' only grab focus for chat edit box on non-touch platforms (i.e. platforms that ought to have a physical keyboard)
                ChatBoxWidget_takeFocus pChatBox
            End If
        End If
    End If

    ' console message history handling
    Do While getNumberConsoleMessages() > getConsoleLineInfo()
        removeTopConsoleMessage
    Loop
    updateConsoleMessages                               ' run the chatbox

    ' widget handling

    If isHeadless Then
        WzTitleUI_run = TITLECODE_CONTINUE ' shortcut everything after this in headless mode
        Exit Function
    End If

    ' Map or player selection is open
    If multiRequestUp Then
        n = widgRunScreen(psRScreen, triggers)
        If n = 0 Then id = 0 Else id = triggers(0).widget.id ' Just use first click here, since the next click could be on another menu.

        isHoverPreview = False
        If runMultiRequester(id, id, sTemp, mapData, isHoverPreview) Then
            Select Case id
            Case MULTIOP_PNAME
                sPlayer_new = sstrcpy_cut(sTemp, 128)
                removeWildcards sPlayer_new

                If loadMultiStats(sPlayer_new, playerStats) Then
                    sPlayer = sPlayer_new
                    widgSetString psWScreen, MULTIOP_PNAME, sTemp
                    printConsoleNameChange NetPlay_players_name(selectedPlayer), sPlayer
                    NETchangePlayerName selectedPlayer, sPlayer
                    setMultiStats selectedPlayer, playerStats, False
                    setMultiStats selectedPlayer, playerStats, True
                    netPlayersUpdated = True
                    If NetPlay_isHost And NetPlay_bComms Then
                        sendOptions
                        NETsetLobbyOptField NetPlay_players_name(selectedPlayer), NET_LOBBY_OPT_FIELD_HOSTNAME
                    End If
                Else
                    debug_ 6556, LOG_ERROR, "run", "Failed to load or create player profile: %s", sPlayer_new
                End If
            Case MULTIOP_MAP
                If isHoverPreview Then
                    oldGameMap = sstrcpy_cut(game.map, 128)
                    oldGameHash = game.hash
                    oldMaxPlayers = game.maxPlayers
                    oldGameIsMapMod = game.isMapMod
                    oldGameIsRandom = game.isRandom

                    game.map = sstrcpy_cut(mapData.pName, 128)
                    game.hash = levGetFileHash(mapData)
                    game.maxPlayers = mapData.players
                    game.isMapMod = CheckForMod(mapData.realFileName)
                    game.isRandom = CheckForRandom(mapData.realFileName, mapData.apDataFiles(0))
                    loadMapPreview False

                    ' Change game info to match the previous selection if hover preview was displayed
                    game.map = oldGameMap
                    game.hash = oldGameHash
                    game.maxPlayers = oldMaxPlayers
                    game.isMapMod = oldGameIsMapMod
                    game.isRandom = oldGameIsRandom
                    GoTo endMap_
                End If

                If NetPlay_bComms And NetPlay_isHost Then
                    numHumans = NET_numHumanPlayers()
                    If numHumans > mapData.players Then
                        displayRoomSystemMessage "Cannot change to a map with too few slots for all players."
                        GoTo endMap_
                    End If
                    If mapData.players < game.maxPlayers Then
                        displayRoomSystemMessage "Cannot change to a map with fewer slots."
                        GoTo endMap_
                    End If
                    If Not canChangeMapOrRandomize() Then
                        GoTo endMap_
                    End If
                End If

                oldMaxPlayers = game.maxPlayers
                updateMapSettings mapData

                loadMapPreview False
                loadMapChallengeAndPlayerSettings
                If Len(mapData.pName) <> 0 Then mapName = mapData.pName Else mapName = "n/a"
                debug_ 6611, LOG_INFO, "run", "Switching map: %s (builtin: %d)", mapName, -CLng(builtInMap)

                'Reset player slots if it's a smaller map.
                If NetPlay_isHost And NetPlay_bComms And oldMaxPlayers > game.maxPlayers Then
                    resetPlayerPositions
                    repositionHumanSlots

                    nHumans = NET_getHumanPlayers(humans)
                    playerInc = 0

                    slotInc = 0
                    Do While slotInc < game.maxPlayers And playerInc < nHumans
                        changePosition humans(playerInc), slotInc, realSelectedPlayer
                        playerInc = playerInc + 1
                        slotInc = slotInc + 1
                    Loop
                End If

                addGameOptions

                If NetPlay_isHost And NetPlay_bComms Then
                    sendOptions
                    NETsetLobbyOptField game.map, NET_LOBBY_OPT_FIELD_MAPNAME
                End If

                wz_command_interface_output_room_status_json True
endMap_:
            Case Else
                loadMapPreview False  ' Restore the preview of the old map.
            End Select
            If Not isHoverPreview Then
                addPlayerBox ingame.side <> InGameSide_HOST_OR_SINGLEPLAYER
            End If
        End If
    ' Map/Player selection (multi-requester) is closed
    Else
        If hideTime <> 0 Then
            ' we abort the 'hidetime' on press of a mouse button.
            If gameTime - hideTime < MAP_PREVIEW_DISPLAY_TIME And Not mousePressed(MOUSE_LMB) And Not mousePressed(MOUSE_RMB) Then
                WzTitleUI_run = TITLECODE_CONTINUE
                Exit Function
            End If
            inputLoseFocus  ' remove the mousepress from the input stream.
            hideTime = 0
        Else
            n = widgRunScreen(psWScreen, triggers)
            If n = 0 Then id = 0 Else id = triggers(0).widget.id ' Just use first click here, since the next click could be on another menu.

            If n <> 0 And (Not multiplayIsStartingGame() Or id = CON_CANCEL) Then
                processMultiopWidgets id
            End If
        End If
    End If

    widgDisplayScreen psWScreen                                 ' show the widgets currently running

    If multiRequestUp Then
        widgDisplayScreen psRScreen                             ' show the Requester running
    End If

    If CancelPressed() Then
        processMultiopWidgets CON_CANCEL  ' "Press" the cancel button to clean up net connections and stuff.
    End If
    If Not NetPlay_isHostAlive And ingame.side = InGameSide_MULTIPLAYER_CLIENT Then
        cancelOrDismissVoteNotifications
        cancelOrDismissNotificationsWithTag LOBBY_DISABLED_TAG
        cancelOrDismissNotificationIfTagHasPrefix SLOTTYPE_TAG_PREFIX ' (the lambda: tag.rfind(SLOTTYPE_TAG_PREFIX, 0) == 0)
        stopJoining WzMsgBoxTitleUI_new("Connection lost:", "The host has quit.", parent), ERROR_HOSTDROPPED
        pie_LoadBackDrop SCREEN_RANDOMBDROP
    End If

    WzTitleUI_run = TITLECODE_CONTINUE
End Function

Private Sub WzTitleUI_screenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ' NOTE: To properly support resizing the inline overlay screen based on underlying screen layer recalculations
    ' frontendScreenSizeDidChange() should be called after intScreenSizeDidChange() in gameScreenSizeDidChange()
End Sub

Private Sub WzTitleUI_start()
    Dim bReenter As Boolean, noLocks As MultiplayOptionsLocked, playerStats As PLAYERSTATS_, hostOrSingle As Boolean
    Dim newRootFrm As WIDGET, psWeakTitleUI As Long, i As Long, psBackdrop As WIDGET
    bReenter = performedFirstStart
    performedFirstStart = True
    netPlayersUpdated = True

    Set psBackdrop = addBackdrop()
    psBackdrop.setCalcLayout WzFunc_new(FN_backdrop_calcLayout)
    addTopForm True

    ' Entering the first time
    If Not bReenter Then
        ' Initialize the inline chooser overlay screen
        Set psInlineChooserOverlayScreen = W_SCREEN_make()
        Set newRootFrm = W_FULLSCREENOVERLAY_CLICKFORM_make(MULTIOP_INLINE_OVERLAY_ROOT_FRM)
        psWeakTitleUI = ObjPtr(Me)
        Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_inlineOverlay_onClicked, psWeakTitleUI)
        Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onCancelPressed = W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc
        psInlineChooserOverlayScreen.psForm.attach newRootFrm

        currentMultiOptionsTitleUI = ObjPtr(Me)
        bRequestedSelfMoveToPlayers = False
        ' (bHostRequestedMoveToPlayers.resize(MAX_CONNECTED_PLAYERS, false): the array has that size)
        For i = 0 To playerRows_size - 1        ' playerRows.clear()
            Set playerRows(i) = Nothing
        Next
        playerRows_size = 0
        Erase playerRows
        initKnownPlayers
        resetPlayerConfiguration True
        locked = noLocks            ' memset(&locked, 0, sizeof(locked))
        If ingame.side = InGameSide_HOST_OR_SINGLEPLAYER Then
            spectatorHost = False
            game.blindMode = BLIND_MODE_NONE
        End If
        defaultOpenSpectatorSlots = war_getMPopenSpectatorSlots()
        If Not loadMapChallengeAndPlayerSettings(True) Then
            If challengeActive Then
                changeTitleUI WzMsgBoxTitleUI_new("Failed to load challenge:", "Failed to load the challenge's map or config", parent)
                Exit Sub
            End If
            If getHostLaunch() = HostLaunch_Autohost Then
                changeTitleUI WzMsgBoxTitleUI_new("Failed to process autohost config:", va("Failed to load the autohost map or config from: %s", wz_skirmish_test()), parent)
                resetHostLaunch ' Don't load the autohost file on subsequent hosts
                Exit Sub
            End If
            ' otherwise, treat as non-fatal...
        End If
        game.isMapMod = False
        game.isRandom = False
        game.mapHasScavengers = True ' FIXME, should default to false

        inlineChooserUp = -1
        aiChooserUp = -1
        difficultyChooserUp = -1
        has_playerSlotSwapChooserUp = False

        ' (the lambda: the identifiers that start with notificationIdentifierPrefix)
        removeNotificationPreferencesIfHasPrefix SLOTTYPE_NOTIFICATION_ID_PREFIX

        ingame.localOptionsReceived = False

        loadMultiStats sPlayer, playerStats
        setMultiStats selectedPlayer, playerStats, True ' just set this locally so that the identity is cached, if changing name

        ' Entering the first time with challenge, immediately start the host
        If challengeActive Then
            If Not startHost() Then
                debug_ 6901, LOG_ERROR, "start", "Failed to host the challenge."
                Exit Sub
            End If
        End If
    End If

    loadMapPreview False

    hostOrSingle = ingame.side = InGameSide_HOST_OR_SINGLEPLAYER
    ' Re-entering or entering without a challenge
    If bReenter Or Not challengeActive Then
        addPlayerBox False
        addGameOptions
        addChatBox bReenter

        If hostOrSingle Then
            printHostHelpMessagesToConsole
        End If
    End If
    If Not bReenter And challengeActive And hostOrSingle Then
        printHostHelpMessagesToConsole ' Print the challenge text the first time
    End If
    If Not bReenter And game.blindMode <> BLIND_MODE_NONE And ingame.side = InGameSide_MULTIPLAYER_CLIENT Then
        printBlindModeHelpMessagesToConsole
    End If

    ' Reset structure limits if we are entering the first time or if we have a challenge
    If Not bReenter Or challengeActive Then
        resetLimits
        ingame.flags = ingame.flags And Not MPFLAGS_FORCELIMITS
        createLimitSet ' should effectively just free it
        updateStructureDisabledFlags
        updateGameOptions
    End If

    If autogame_enabled() Or getHostLaunch() = HostLaunch_Autohost Then
        If Not ingame.localJoiningInProgress Then
            startHost
        End If
        If Not getHostLaunchStartNotReady() Then
            sendReadyRequest selectedPlayer, True
        End If
        If getHostLaunch() = HostLaunch_Skirmish Then
            startMultiplayerGame
            ' reset flag in case people dropped/quit on join screen
            NETsetPlayerConnectionStatus CONNECTIONSTATUS_NORMAL, NET_ALL_PLAYERS
        End If
    End If
End Sub

Public Function getParentTitleUI() As WzTitleUI
    Set getParentTitleUI = parent
End Function
