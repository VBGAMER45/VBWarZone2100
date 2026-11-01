Attribute VB_Name = "teamstrategy_c"
' src/hci/teamstrategy.cpp, teamstrategy.h, whole: the team strategy plans (sent to the human teammates with
' NET_TEAM_STRATEGY) and the team strategy view (the pre-game screen's team planning table; multimenu's in-game tab
' is cluster G's, not ported). Its widget classes' fields are the data classes named as the classes.
' (VB) class WzTeamStrategyColumnImagesManager is WzTeamStratImagesManager.cls (a plain class: its textures are
' released with it); class WzTeamStrategyColumnImageRef is the UDT WzTeamStrategyColumnImageRef. The
' std::weak_ptr<TeamStrategyView> of the buttons is the view's address (WIDGET_fromPtr), as in the other ported
' screens. stats.cpp's loadImageForWeapSubclass is here (stats_c has not ported it). aiInformTeamOfStrategy's
' unordered_maps are key and value arrays. Translations (_()) are not ported: the English text is used.
' (VB) widget class WzTeamStrategyIconColumnHeader : WIDGET
' (VB) widget class WzTeamStrategyButton : W_BUTTON
' (VB) virtual WzTeamStrategyButton.clickPrimary
' (VB) virtual WzTeamStrategyButton.clickSecondary
' (VB) virtual WzTeamStrategyButton.currentState
' (VB) widget class WzTeamStrategyWeaponButton : WzTeamStrategyButton
' (VB) widget class WzTeamStrategyUnitTypeButton : WzTeamStrategyButton
' (VB) widget class TeamStrategyView : W_FORM
' (VB) widget class PlayerNameLabel : W_LABEL
Option Explicit

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' teamstrategy.h: enum class WzStrategyPlanningState : uint8_t
Public Enum WzStrategyPlanningState
    WzStrategyPlanningState_NotSpecified
    WzStrategyPlanningState_Planned
    WzStrategyPlanningState_HeavilyPlanned
    WzStrategyPlanningState_Never
End Enum

' teamstrategy.h: enum class WzStrategyPlanningUnitTypes : uint8_t
Public Enum WzStrategyPlanningUnitTypes
    WzStrategyPlanningUnitTypes_Tanks
    WzStrategyPlanningUnitTypes_VTOL
    WzStrategyPlanningUnitTypes_Cyborg
End Enum
Public Const WzStrategyPlanningUnitTypes_NUM_TYPES As Long = 3

' class WzTeamStrategyColumnImageRef (optional<WEAPON_SUBCLASS> weaponSubclass, optional<WzStrategyPlanningUnitTypes>
' unitType, std::shared_ptr<WzTeamStrategyColumnImagesManager> sharedImageManager)
Public Type WzTeamStrategyColumnImageRef
    has_weaponSubclass As Boolean
    weaponSubclass As WEAPON_SUBCLASS
    has_unitType As Boolean
    unitType As WzStrategyPlanningUnitTypes
    sharedImageManager As WzTeamStratImagesManager
End Type

' the lambdas (WzFunc: teamstrategy_c_fn)
Public Const FN_teamStrat_noTeammatesLabel_calcLayout As Long = 17010  ' result->noTeammatesLabel->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))
Public Const FN_teamStrat_table_calcLayout As Long = 17011             ' result->table->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))

' MARK: - Globals and helper functions

Private Const COL_PADDING_X As Long = 2
Private Const MAX_TEAMSTRAT_COLUMN_HEADERS_HEIGHT As Long = 120

Private Type WzStrategyPlanningDetails
    state As WzStrategyPlanningState
    localPlayerHasSeen As Boolean       ' = true
End Type

Private Type WzPlayerStrategyPlan
    weaponPlans(0 To WSC_NUM_WEAPON_SUBCLASSES - 1) As WzStrategyPlanningDetails   ' std::array<WzStrategyPlanningDetails, WSC_NUM_WEAPON_SUBCLASSES> WeaponPlans
    unitPlans(0 To WzStrategyPlanningUnitTypes_NUM_TYPES - 1) As WzStrategyPlanningDetails  ' std::array<WzStrategyPlanningDetails, WzStrategyPlanningUnitTypes_NUM_TYPES> UnitPlans
    hasBeenSet As Boolean               ' = false
End Type

' NOTE: Only filled out for teammates, as only teammates send their information
' (VB) the details' localPlayerHasSeen start False here (true in C++) until strategyPlansInit: nothing reads it
Private playerStrategyPlans(0 To MAX_PLAYERS - 1) As WzPlayerStrategyPlan

Private Const WZ_TEAMSTRAT_COLUMN_HEADER_INTERNAL_PADDING As Long = 4
Private Const WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_Y As Long = 4
Private Const WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_X As Long = 1
Private Const WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE As Long = 16
Private Const MAX_PLAYER_NAME_COLUMN_WIDTH As Long = 120

Private Function to_WzStrategyPlanningState(ByVal value As Long, output As WzStrategyPlanningState) As Boolean
    If value >= WzStrategyPlanningState_NotSpecified And value <= WzStrategyPlanningState_Never Then
        output = value
        to_WzStrategyPlanningState = True
        Exit Function
    End If
    to_WzStrategyPlanningState = False
End Function

Private Function cycleWzStrategyPlanningState(ByVal state As WzStrategyPlanningState, Optional ByVal directionForward As Boolean = True) As WzStrategyPlanningState
    Select Case state
    Case WzStrategyPlanningState_Never
        If directionForward Then cycleWzStrategyPlanningState = WzStrategyPlanningState_NotSpecified Else cycleWzStrategyPlanningState = WzStrategyPlanningState_HeavilyPlanned
        Exit Function
    Case WzStrategyPlanningState_NotSpecified
        If directionForward Then cycleWzStrategyPlanningState = WzStrategyPlanningState_Planned Else cycleWzStrategyPlanningState = WzStrategyPlanningState_Never
        Exit Function
    Case WzStrategyPlanningState_Planned
        If directionForward Then cycleWzStrategyPlanningState = WzStrategyPlanningState_HeavilyPlanned Else cycleWzStrategyPlanningState = WzStrategyPlanningState_NotSpecified
        Exit Function
    Case WzStrategyPlanningState_HeavilyPlanned
        If directionForward Then cycleWzStrategyPlanningState = WzStrategyPlanningState_Never Else cycleWzStrategyPlanningState = WzStrategyPlanningState_Planned
        Exit Function
    End Select
    cycleWzStrategyPlanningState = WzStrategyPlanningState_NotSpecified ' silence warning
End Function

Private Function getUnitTypeDisplayName(ByVal unitType As WzStrategyPlanningUnitTypes) As String
    Select Case unitType
    Case WzStrategyPlanningUnitTypes_Tanks
        getUnitTypeDisplayName = "Tanks"
        Exit Function
    Case WzStrategyPlanningUnitTypes_VTOL
        getUnitTypeDisplayName = "VTOL"
        Exit Function
    Case WzStrategyPlanningUnitTypes_Cyborg
        getUnitTypeDisplayName = "Cyborgs"
        Exit Function
    End Select
    getUnitTypeDisplayName = "" ' silence compiler warning
End Function

' void WzStrategyPlanningDetails::reset()
Private Sub WzStrategyPlanningDetails_reset(details As WzStrategyPlanningDetails)
    details.state = WzStrategyPlanningState_NotSpecified
    details.localPlayerHasSeen = True
End Sub

' (VB) a Byte array's size (0 when unallocated)
Private Function byteArraySize(b() As Byte) As Long
    On Error GoTo empty_
    byteArraySize = UBound(b) - LBound(b) + 1
    Exit Function
empty_:
    byteArraySize = 0
End Function

' MARK: -

Public Function checkedGetPlayerTeam(ByVal i As Long) As Long
    If alliancesSetTeamsBeforeGame(game.alliance) Then
        checkedGetPlayerTeam = NetPlay_players_team(i)
    Else
        checkedGetPlayerTeam = i
    End If
End Function

Private Sub sendStrategyPlanUpdate(ByVal forPlayer As Long)
    Dim weaponStates() As Byte, unitStates() As Byte, i As Long, sendingPlayerTeam As Long, player As Long
    Dim w As MessageWriter, sender As Long
    If Not ULT(forPlayer, MAX_PLAYERS) Then ASSERT_FAILURE "forPlayer < playerStrategyPlans.size()", "src/hci/teamstrategy.cpp:135", "sendStrategyPlanUpdate", "Unexpected player: %u", forPlayer: Exit Sub
    If Not myResponsibility(forPlayer) Then ASSERT_FAILURE "myResponsibility(forPlayer)", "src/hci/teamstrategy.cpp:136", "sendStrategyPlanUpdate", "We are not responsible for player: %u", forPlayer: Exit Sub

    ' build parts of the messages
    ReDim weaponStates(0 To WSC_NUM_WEAPON_SUBCLASSES - 1)
    For i = 0 To WSC_NUM_WEAPON_SUBCLASSES - 1
        weaponStates(i) = playerStrategyPlans(forPlayer).weaponPlans(i).state
    Next
    ReDim unitStates(0 To WzStrategyPlanningUnitTypes_NUM_TYPES - 1)
    For i = 0 To WzStrategyPlanningUnitTypes_NUM_TYPES - 1
        unitStates(i) = playerStrategyPlans(forPlayer).unitPlans(i).state
    Next

    ' send updates to every (human) team member
    sendingPlayerTeam = checkedGetPlayerTeam(forPlayer)
    For player = 0 To MIN_(game.maxPlayers, MAX_PLAYERS) - 1
        If player = forPlayer Then
            GoTo continue_
        End If
        If checkedGetPlayerTeam(player) <> sendingPlayerTeam Then
            GoTo continue_
        End If
        ' Only send to active (connected) human players - *NOT* AIs (which would effectively just have to send this to the host)
        If isHumanPlayer(player) Then
            Set w = NETbeginEncodeSecured(NETnetQueue(player), NET_TEAM_STRATEGY)
            If Not w Is Nothing Then
                sender = forPlayer
                NETuint32_t w, sender
                NETbytes w, weaponStates
                NETbytes w, unitStates
                NETend w
            End If
        End If
continue_:
    Next
End Sub

Public Function recvStrategyPlanUpdate(ByVal queue As Long) As Boolean
    Dim sender As Long, weaponStates() As Byte, unitStates() As Byte, r As MessageReader, selectedPlayerTeam As Long
    Dim i As Long, newState As WzStrategyPlanningState
    sender = MAX_CONNECTED_PLAYERS

    Set r = NETbeginDecodeSecured(queue, NET_TEAM_STRATEGY)
    If r Is Nothing Then
        recvStrategyPlanUpdate = False
        Exit Function
    End If
    NETuint32_t r, sender
    NETbytes r, weaponStates
    NETbytes r, unitStates
    NETend r

    If whosResponsible(sender) <> NETQUEUE_index(queue) Then
        recvStrategyPlanUpdate = False
        Exit Function
    End If

    If Not ULT(sender, MAX_PLAYERS) Then
        recvStrategyPlanUpdate = False
        Exit Function
    End If
    If Not NetPlay_players_allocated(sender) Then
        recvStrategyPlanUpdate = False
        Exit Function
    End If

    selectedPlayerTeam = checkedGetPlayerTeam(selectedPlayer)
    If checkedGetPlayerTeam(sender) <> selectedPlayerTeam Then
        recvStrategyPlanUpdate = False
        Exit Function
    End If

    ' update any changed data
    For i = 0 To MIN_(WSC_NUM_WEAPON_SUBCLASSES, byteArraySize(weaponStates)) - 1
        If Not to_WzStrategyPlanningState(weaponStates(i), newState) Then
            GoTo continue1
        End If
        If newState <> playerStrategyPlans(sender).weaponPlans(i).state Then
            playerStrategyPlans(sender).weaponPlans(i).state = newState
            playerStrategyPlans(sender).weaponPlans(i).localPlayerHasSeen = False
        End If
continue1:
    Next

    For i = 0 To MIN_(WzStrategyPlanningUnitTypes_NUM_TYPES, byteArraySize(unitStates)) - 1
        If Not to_WzStrategyPlanningState(unitStates(i), newState) Then
            GoTo continue2
        End If
        If newState <> playerStrategyPlans(sender).unitPlans(i).state Then
            playerStrategyPlans(sender).unitPlans(i).state = newState
            playerStrategyPlans(sender).unitPlans(i).localPlayerHasSeen = False
        End If
continue2:
    Next

    playerStrategyPlans(sender).hasBeenSet = True

    recvStrategyPlanUpdate = True
End Function

Public Sub strategyPlansInit()
    Dim i As Long
    For i = 0 To MAX_PLAYERS - 1
        WzPlayerStrategyPlan_reset i
    Next
End Sub

' void WzPlayerStrategyPlan::reset() (of playerStrategyPlans[player])
Private Sub WzPlayerStrategyPlan_reset(ByVal player As Long)
    Dim i As Long
    For i = 0 To WSC_NUM_WEAPON_SUBCLASSES - 1
        WzStrategyPlanningDetails_reset playerStrategyPlans(player).weaponPlans(i)
    Next
    For i = 0 To WzStrategyPlanningUnitTypes_NUM_TYPES - 1
        WzStrategyPlanningDetails_reset playerStrategyPlans(player).unitPlans(i)
    Next
    playerStrategyPlans(player).hasBeenSet = False
End Sub

' MARK: -

' WzTeamStrategyColumnImageRef(WEAPON_SUBCLASS subClass, const std::shared_ptr<WzTeamStrategyColumnImagesManager>&)
' (WzTeamStrategyColumnImagesManager::getWeaponSubclassImageRef)
Public Function WzTeamStrategyColumnImageRef_weapon(ByVal subClass As WEAPON_SUBCLASS, ByVal sharedImageManager As WzTeamStratImagesManager) As WzTeamStrategyColumnImageRef
    WzTeamStrategyColumnImageRef_weapon.has_weaponSubclass = True
    WzTeamStrategyColumnImageRef_weapon.weaponSubclass = subClass
    Set WzTeamStrategyColumnImageRef_weapon.sharedImageManager = sharedImageManager
End Function

' WzTeamStrategyColumnImageRef(WzStrategyPlanningUnitTypes unitType, const std::shared_ptr<...>&)
' (WzTeamStrategyColumnImagesManager::getUnitTypeImageRef)
Public Function WzTeamStrategyColumnImageRef_unitType(ByVal unitType As WzStrategyPlanningUnitTypes, ByVal sharedImageManager As WzTeamStratImagesManager) As WzTeamStrategyColumnImageRef
    WzTeamStrategyColumnImageRef_unitType.has_unitType = True
    WzTeamStrategyColumnImageRef_unitType.unitType = unitType
    Set WzTeamStrategyColumnImageRef_unitType.sharedImageManager = sharedImageManager
End Function

Public Function WzTeamStrategyColumnImageRef_getImageTexture(ref As WzTeamStrategyColumnImageRef) As gfx_api_texture
    If ref.has_weaponSubclass Then
        Set WzTeamStrategyColumnImageRef_getImageTexture = ref.sharedImageManager.getWeaponSubclassImageTexture(ref.weaponSubclass)
        Exit Function
    ElseIf ref.has_unitType Then
        Set WzTeamStrategyColumnImageRef_getImageTexture = ref.sharedImageManager.getUnitTypeImageTexture(ref.unitType)
        Exit Function
    End If
    debug_ 380, LOG_ERROR, "getImageTexture", "Invalid reference"
    Set WzTeamStrategyColumnImageRef_getImageTexture = Nothing
End Function

' stats.cpp: gfx_api::texture* loadImageForWeapSubclass(WEAPON_SUBCLASS subClass)
Public Function loadImageForWeapSubclass(ByVal subClass As WEAPON_SUBCLASS) As gfx_api_texture
    Dim imagePath As String
    Select Case subClass
    Case WSC_MGUN
        imagePath = "images/intfac/wsc_mgun.png"
    Case WSC_CANNON
        imagePath = "images/intfac/wsc_cannon.png"
    Case WSC_MORTARS
        imagePath = "images/intfac/wsc_mortars.png"
    Case WSC_MISSILE
        imagePath = "images/intfac/wsc_missile.png"
    Case WSC_ROCKET
        imagePath = "images/intfac/wsc_rocket.png"
    Case WSC_ENERGY
        imagePath = "images/intfac/wsc_energy.png"
    Case WSC_GAUSS
        imagePath = "images/intfac/wsc_gauss.png"
    Case WSC_FLAME
        imagePath = "images/intfac/wsc_flame.png"
    'case WSC_CLOSECOMBAT:
    Case WSC_HOWITZERS
        imagePath = "images/intfac/wsc_howitzers.png"
    Case WSC_ELECTRONIC
        imagePath = "images/intfac/wsc_electronic.png"
    Case WSC_AAGUN
        imagePath = "images/intfac/wsc_aagun.png"
    Case WSC_SLOWMISSILE
        imagePath = "images/intfac/wsc_slowmissile.png"
    Case WSC_SLOWROCKET
        imagePath = "images/intfac/wsc_slowrocket.png"
    Case WSC_LAS_SAT
        imagePath = "images/intfac/wsc_las_sat.png"
    Case WSC_BOMB
        imagePath = "images/intfac/wsc_bomb.png"
    Case WSC_COMMAND
        imagePath = "images/intfac/wsc_command.png"
    Case WSC_EMP
        imagePath = "images/intfac/wsc_emp.png"
    Case WSC_NUM_WEAPON_SUBCLASSES  ' The number of enumerators in this enum.
    End Select
    If Not (Len(imagePath) <> 0) Then ASSERT_FAILURE "imagePath != nullptr", "src/stats.cpp:1935", "loadImageForWeapSubclass", "No image path": Exit Function
    ' WZ_Notification_Image img(imagePath); return img.loadImageToTexture();
    Set loadImageForWeapSubclass = gfx_api_context_loadTextureFromFile(imagePath, texture_type_user_interface)
End Function

Public Function loadImageForUnitType(ByVal unitType As WzStrategyPlanningUnitTypes) As gfx_api_texture
    Dim imagePath As String
    Select Case unitType
    Case WzStrategyPlanningUnitTypes_Tanks
        imagePath = "images/intfac/unittype_tank.png"
    Case WzStrategyPlanningUnitTypes_VTOL
        imagePath = "images/intfac/unittype_vtol.png"
    Case WzStrategyPlanningUnitTypes_Cyborg
        imagePath = "images/intfac/unittype_cyborg.png"
    End Select
    If Not (Len(imagePath) <> 0) Then ASSERT_FAILURE "imagePath != nullptr", "src/hci/teamstrategy.cpp:420", "loadImageForUnitType", "No image path": Exit Function
    ' WZ_Notification_Image img(imagePath); return img.loadImageToTexture();
    Set loadImageForUnitType = gfx_api_context_loadTextureFromFile(imagePath, texture_type_user_interface)
End Function

' MARK: - WzTeamStrategyIconColumnHeader


' WzTeamStrategyIconColumnHeader(const WzTeamStrategyColumnImageRef& columnImageRef)
Public Sub WzTeamStrategyIconColumnHeader_construct(ByVal self As WIDGET, columnImageRef As WzTeamStrategyColumnImageRef)
    WIDGET_construct self
    self.kind = K_WzTeamStrategyIconColumnHeader ' (VB) virtual calls from here on reach WzTeamStrategyIconColumnHeader's overrides
    WzTeamStrategyIconColumnHeader_(self).columnImageRef = columnImageRef
End Sub

Public Function WzTeamStrategyIconColumnHeader_make(ByVal text As String, columnImageRef As WzTeamStrategyColumnImageRef, Optional ByVal disabled As Boolean = False) As WIDGET
    Dim result As WIDGET, d As WzTeamStrategyIconColumnHeader
    ' (class make_shared_enabler)
    Set result = WIDGET_alloc(K_WzTeamStrategyIconColumnHeader)
    WzTeamStrategyIconColumnHeader_construct result, columnImageRef
    Set d = WzTeamStrategyIconColumnHeader_(result)

    d.columnText.setText text, font_regular
    d.disabledColumn = disabled

    Set WzTeamStrategyIconColumnHeader_make = result
End Function

Public Sub WzTeamStrategyIconColumnHeader_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzTeamStrategyIconColumnHeader, x0 As Long, y0 As Long, iconColor As Long, pImageTexture As gfx_api_texture
    Dim imageX0 As Long, imageY0 As Long, textX0 As Long, textY0 As Long, maxTextDisplayableWidth As Long
    Dim willTruncate As Boolean, textColor As Long, ref As WzTeamStrategyColumnImageRef
    Set d = WzTeamStrategyIconColumnHeader_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    ' Draw the column image (at the bottom of the widget
    If Not d.disabledColumn Then iconColor = pal_RGBA(255, 255, 255, 200) Else iconColor = pal_RGBA(200, 200, 200, 100)
    ref = d.columnImageRef
    Set pImageTexture = WzTeamStrategyColumnImageRef_getImageTexture(ref)
    If Not pImageTexture Is Nothing Then
        imageX0 = x0 + (self.width() - WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE) \ 2
        imageY0 = y0 + (self.height() - (WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_Y + WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE))
        iV_DrawImageAnisotropic pImageTexture, imageX0, imageY0, 0!, 0!, WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE, WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE, 0!, iconColor
    End If
    ' Draw the text, sideways
    textX0 = ((self.width() - d.columnText.lineSize()) \ 2) + (-d.columnText.aboveBase())
    textY0 = self.height() - (WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_Y + WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE + WZ_TEAMSTRAT_COLUMN_HEADER_INTERNAL_PADDING)
    maxTextDisplayableWidth = textY0 - WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_Y
    willTruncate = maxTextDisplayableWidth < d.columnText.width()
    If willTruncate Then
        maxTextDisplayableWidth = maxTextDisplayableWidth - (iV_GetEllipsisWidth(d.columnText.getFontID()) + 2)
    End If
    If Not d.disabledColumn Then textColor = psPalette(WZCOL_TEXT_DARK) Else textColor = psPalette(WZCOL_GREY)
    d.columnText.render x0 + textX0, y0 + textY0, textColor, 270!, maxTextDisplayableWidth

    If willTruncate Then
        ' Render ellipsis, also sideways
        iV_DrawEllipsis d.columnText.getFontID(), x0 + textX0, (y0 + textY0) - (maxTextDisplayableWidth + 2), textColor, 270!
    End If
End Sub

Public Function WzTeamStrategyIconColumnHeader_idealWidth(ByVal self As WIDGET) As Long
    WzTeamStrategyIconColumnHeader_idealWidth = WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_X + WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE + WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_X
End Function

Public Function WzTeamStrategyIconColumnHeader_idealHeight(ByVal self As WIDGET) As Long
    WzTeamStrategyIconColumnHeader_idealHeight = WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_Y + WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE + WZ_TEAMSTRAT_COLUMN_HEADER_INTERNAL_PADDING + WzTeamStrategyIconColumnHeader_(self).columnText.width() + WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_Y
End Function

Public Sub WzTeamStrategyIconColumnHeader_setTip(ByVal self As WIDGET, ByVal string_ As String)
    WzTeamStrategyIconColumnHeader_(self).pTip = string_
End Sub

Public Function WzTeamStrategyIconColumnHeader_getTip(ByVal self As WIDGET) As String
    WzTeamStrategyIconColumnHeader_getTip = WzTeamStrategyIconColumnHeader_(self).pTip
End Function

' MARK: - WzTeamStrategyButton

' WzTeamStrategyButton(const std::shared_ptr<TeamStrategyView>& _parentStrategyView, const WzTeamStrategyColumnImageRef&
' imageRef, int displayWidth, int displayHeight)
Public Sub WzTeamStrategyButton_construct(ByVal self As WIDGET, ByVal parentStrategyView_ As WIDGET, imageRef As WzTeamStrategyColumnImageRef, ByVal displayWidth As Long, ByVal displayHeight As Long)
    Dim d As WzTeamStrategyButton
    W_BUTTON_construct self
    self.kind = K_WzTeamStrategyButton ' (VB) virtual calls from here on reach WzTeamStrategyButton's overrides
    Set d = WzTeamStrategyButton_(self)
    d.parentStrategyView = ObjPtr(parentStrategyView_)
    Set d.sharedImageManager = TeamStrategyView_(parentStrategyView_).columnImagesManager
    d.imageRef = imageRef
    d.displayWidth = displayWidth
    d.displayHeight = displayHeight
    self.style = self.style Or WBUT_SECONDARY
End Sub

' special click-handling functions - subclasses should override
Public Sub WzTeamStrategyButton_clickPrimary(ByVal self As WIDGET)
End Sub

Public Function WzTeamStrategyButton_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean) As Boolean
    WzTeamStrategyButton_clickSecondary = False
End Function

' (pure virtual)
Public Function WzTeamStrategyButton_currentState(ByVal self As WIDGET) As WzStrategyPlanningState
    WzTeamStrategyButton_currentState = WzStrategyPlanningState_NotSpecified
End Function

' override default button click handling
Public Sub WzTeamStrategyButton_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim clickAndReleaseOnThisButton As Boolean
    clickAndReleaseOnThisButton = ((W_BUTTON_(self).state And WBUT_DOWN) <> 0) ' relies on W_BUTTON handling to properly set WBUT_DOWN

    W_BUTTON_released self, psContext, key

    If Not clickAndReleaseOnThisButton Then
        Exit Sub ' do nothing
    End If

    If key = WKEY_PRIMARY Then
        vt_WzTeamStrategyButton_clickPrimary self
    ElseIf key = WKEY_SECONDARY Then
        vt_WzTeamStrategyButton_clickSecondary self, False
    End If
End Sub

Public Function WzTeamStrategyButton_clickHeld(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY) As Boolean
    If key = WKEY_PRIMARY Then
        WzTeamStrategyButton_clickHeld = vt_WzTeamStrategyButton_clickSecondary(self, True)
        Exit Function
    End If
    WzTeamStrategyButton_clickHeld = False
End Function

Public Function WzTeamStrategyButton_idealWidth(ByVal self As WIDGET) As Long
    WzTeamStrategyButton_idealWidth = WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_X + WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE + WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_X
End Function

Public Function WzTeamStrategyButton_idealHeight(ByVal self As WIDGET) As Long
    WzTeamStrategyButton_idealHeight = WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_Y + WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE + WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_Y
End Function

Public Sub WzTeamStrategyButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzTeamStrategyButton, x0 As Long, y0 As Long, strategyState As WzStrategyPlanningState, isDisabled As Boolean
    Dim isHighlight As Boolean, imageX0 As Long, imageY0 As Long, pImageTexture As gfx_api_texture, ref As WzTeamStrategyColumnImageRef
    Set d = WzTeamStrategyButton_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    strategyState = vt_WzTeamStrategyButton_currentState(self)

'   bool isDown = (state & (WBUT_DOWN | WBUT_LOCK | WBUT_CLICKLOCK)) != 0;
    isDisabled = (W_BUTTON_(self).state And WBUT_DISABLE) <> 0
    isHighlight = (W_BUTTON_(self).state And WBUT_HIGHLIGHT) <> 0

    If isHighlight And Not isDisabled Then
        pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), pal_RGBA(0, 0, 255, 50)
    End If

    ' An icon / image
    imageX0 = x0 + (self.width() - d.displayWidth) \ 2
    imageY0 = y0 + (self.height() - d.displayHeight) \ 2
    If strategyState = WzStrategyPlanningState_Planned Or strategyState = WzStrategyPlanningState_HeavilyPlanned Then
        ref = d.imageRef
        Set pImageTexture = WzTeamStrategyColumnImageRef_getImageTexture(ref)
        If Not pImageTexture Is Nothing Then
            iV_DrawImageAnisotropic pImageTexture, imageX0, imageY0, 0!, 0!, d.displayWidth, d.displayHeight, 0!, pal_RGBA(255, 255, 255, 200)
        End If
    End If
    If strategyState = WzStrategyPlanningState_Never Then
        d.sharedImageManager.renderNeverSymbol x0, y0, self.width(), self.height()
    End If
    ' (optional) A full inner border (if state == extra selected)
    If strategyState = WzStrategyPlanningState_HeavilyPlanned Then
        iV_Box imageX0 - 1, imageY0 - 1, imageX0 + d.displayWidth + 1, imageY0 + d.displayHeight + 1, pal_RGBA(252, 132, 3, 200)
    End If
End Sub

' WzTeamStrategyWeaponButton(const std::shared_ptr<TeamStrategyView>& parentStrategyView, const WzTeamStrategyColumnImageRef&
' imageRef, int displayWidth, int displayHeight)
Public Sub WzTeamStrategyWeaponButton_construct(ByVal self As WIDGET, ByVal parentStrategyView_ As WIDGET, imageRef As WzTeamStrategyColumnImageRef, ByVal displayWidth As Long, ByVal displayHeight As Long)
    WzTeamStrategyButton_construct self, parentStrategyView_, imageRef, displayWidth, displayHeight
    self.kind = K_WzTeamStrategyWeaponButton ' (VB) virtual calls from here on reach WzTeamStrategyWeaponButton's overrides
End Sub

Public Sub WzTeamStrategyWeaponButton_clickPrimary(ByVal self As WIDGET)
    WzTeamStrategyWeaponButton_handleClick self, True
End Sub

Public Function WzTeamStrategyWeaponButton_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean) As Boolean
    WzTeamStrategyWeaponButton_clickSecondary = WzTeamStrategyWeaponButton_handleClick(self, False)
End Function

Public Function WzTeamStrategyWeaponButton_currentState(ByVal self As WIDGET) As WzStrategyPlanningState
    Dim d As WzTeamStrategyWeaponButton
    Set d = WzTeamStrategyWeaponButton_(self)
    WzTeamStrategyWeaponButton_currentState = playerStrategyPlans(d.playerIdx).weaponPlans(d.subClass).state
End Function

Private Function WzTeamStrategyWeaponButton_handleClick(ByVal self As WIDGET, ByVal primary As Boolean) As Boolean
    Dim d As WzTeamStrategyWeaponButton, newState As WzStrategyPlanningState, strongParent As WIDGET
    Set d = WzTeamStrategyWeaponButton_(self)
    If d.playerIdx <> selectedPlayer Then
        WzTeamStrategyWeaponButton_handleClick = False
        Exit Function
    End If

    ' Change the actual stored value for this player + weapon subclass
    newState = cycleWzStrategyPlanningState(playerStrategyPlans(d.playerIdx).weaponPlans(d.subClass).state, primary)
    playerStrategyPlans(d.playerIdx).weaponPlans(d.subClass).state = newState
    playerStrategyPlans(d.playerIdx).hasBeenSet = True

    Set strongParent = WIDGET_fromPtr(WzTeamStrategyButton_(self).parentStrategyView)
    If Not strongParent Is Nothing Then
        If TeamStrategyView_getUpdatingPlayer(strongParent) Then
            ' Send the change to other team players
            sendStrategyPlanUpdate selectedPlayer
        End If
    End If

    WzTeamStrategyWeaponButton_handleClick = True
End Function

' WzTeamStrategyUnitTypeButton(const std::shared_ptr<TeamStrategyView>& parentStrategyView, const WzTeamStrategyColumnImageRef&
' imageRef, int displayWidth, int displayHeight)
Public Sub WzTeamStrategyUnitTypeButton_construct(ByVal self As WIDGET, ByVal parentStrategyView_ As WIDGET, imageRef As WzTeamStrategyColumnImageRef, ByVal displayWidth As Long, ByVal displayHeight As Long)
    WzTeamStrategyButton_construct self, parentStrategyView_, imageRef, displayWidth, displayHeight
    self.kind = K_WzTeamStrategyUnitTypeButton ' (VB) virtual calls from here on reach WzTeamStrategyUnitTypeButton's overrides
End Sub

Public Sub WzTeamStrategyUnitTypeButton_clickPrimary(ByVal self As WIDGET)
    WzTeamStrategyUnitTypeButton_handleClick self, True
End Sub

Public Function WzTeamStrategyUnitTypeButton_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean) As Boolean
    WzTeamStrategyUnitTypeButton_clickSecondary = WzTeamStrategyUnitTypeButton_handleClick(self, False)
End Function

Public Function WzTeamStrategyUnitTypeButton_currentState(ByVal self As WIDGET) As WzStrategyPlanningState
    Dim d As WzTeamStrategyUnitTypeButton
    Set d = WzTeamStrategyUnitTypeButton_(self)
    WzTeamStrategyUnitTypeButton_currentState = playerStrategyPlans(d.playerIdx).unitPlans(d.unitType).state
End Function

Private Function WzTeamStrategyUnitTypeButton_handleClick(ByVal self As WIDGET, ByVal primary As Boolean) As Boolean
    Dim d As WzTeamStrategyUnitTypeButton, newState As WzStrategyPlanningState, strongParent As WIDGET
    Set d = WzTeamStrategyUnitTypeButton_(self)
    If d.playerIdx <> selectedPlayer Then
        WzTeamStrategyUnitTypeButton_handleClick = False
        Exit Function
    End If

    ' Change the actual stored value for this player + weapon subclass
    newState = cycleWzStrategyPlanningState(playerStrategyPlans(d.playerIdx).unitPlans(d.unitType).state, primary)
    playerStrategyPlans(d.playerIdx).unitPlans(d.unitType).state = newState
    playerStrategyPlans(d.playerIdx).hasBeenSet = True

    Set strongParent = WIDGET_fromPtr(WzTeamStrategyButton_(self).parentStrategyView)
    If Not strongParent Is Nothing Then
        If TeamStrategyView_getUpdatingPlayer(strongParent) Then
            ' Send the change to other team players
            sendStrategyPlanUpdate selectedPlayer
        End If
    End If

    WzTeamStrategyUnitTypeButton_handleClick = True
End Function

Public Function WzTeamStrategyWeaponButton_make(ByVal playerIdx As Long, ByVal subClass As WEAPON_SUBCLASS, ByVal parentStrategyView_ As WIDGET) As WIDGET
    Dim imageRef As WzTeamStrategyColumnImageRef, result As WIDGET
    imageRef = WzTeamStrategyColumnImageRef_weapon(subClass, TeamStrategyView_(parentStrategyView_).columnImagesManager)

    ' (class make_shared_enabler)
    Set result = WIDGET_alloc(K_WzTeamStrategyWeaponButton)
    WzTeamStrategyWeaponButton_construct result, parentStrategyView_, imageRef, WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE, WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE

    WzTeamStrategyWeaponButton_(result).playerIdx = playerIdx
    WzTeamStrategyWeaponButton_(result).subClass = subClass

    Set WzTeamStrategyWeaponButton_make = result
End Function

Public Function WzTeamStrategyUnitTypeButton_make(ByVal playerIdx As Long, ByVal unitType As WzStrategyPlanningUnitTypes, ByVal parentStrategyView_ As WIDGET) As WIDGET
    Dim imageRef As WzTeamStrategyColumnImageRef, result As WIDGET
    imageRef = WzTeamStrategyColumnImageRef_unitType(unitType, TeamStrategyView_(parentStrategyView_).columnImagesManager)

    ' (class make_shared_enabler)
    Set result = WIDGET_alloc(K_WzTeamStrategyUnitTypeButton)
    WzTeamStrategyUnitTypeButton_construct result, parentStrategyView_, imageRef, WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE, WZ_TEAMSTRAT_COLUMN_HEADER_IMAGE_DISPLAY_SIZE

    WzTeamStrategyUnitTypeButton_(result).playerIdx = playerIdx
    WzTeamStrategyUnitTypeButton_(result).unitType = unitType

    Set WzTeamStrategyUnitTypeButton_make = result
End Function

' MARK: - TeamStrategyView


' TeamStrategyView(): columnImagesManager(std::make_shared<WzTeamStrategyColumnImagesManager>())
Public Sub TeamStrategyView_construct(ByVal self As WIDGET)
    W_FORM_construct self
    self.kind = K_TeamStrategyView ' (VB) virtual calls from here on reach TeamStrategyView's overrides
    Set TeamStrategyView_(self).columnImagesManager = New WzTeamStratImagesManager
End Sub

' (~TeamStrategyView(): empty)

Public Function TeamStrategyView_getUpdatingPlayer(ByVal self As WIDGET) As Boolean
    TeamStrategyView_getUpdatingPlayer = TeamStrategyView_(self).updatingPlayers
End Function

Public Sub TeamStrategyView_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As TeamStrategyView, x0 As Long, y0 As Long
    Set d = TeamStrategyView_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    If SHRU(d.backgroundColor, 24) <> 0 Then   ' !backgroundColor.isTransparent()
        pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), d.backgroundColor
    End If
End Sub

Public Sub TeamStrategyView_geometryChanged(ByVal self As WIDGET)
    Dim d As TeamStrategyView
    Set d = TeamStrategyView_(self)
    If Not d.noTeammatesLabel Is Nothing Then
        d.noTeammatesLabel.callCalcLayout
    End If
    If Not d.table Is Nothing Then
        d.table.callCalcLayout
    End If
End Sub

Private Sub TeamStrategyView_updateNoTeamLabelContents(ByVal self As WIDGET)
    Dim d As TeamStrategyView, noTeammatesStr As String
    Set d = TeamStrategyView_(self)
    noTeammatesStr = "No teammates."
    If Not ULT(GAME_TICKS_PER_UPDATE, gameTime) Then   ' gameTime <= GAME_TICKS_PER_UPDATE
        noTeammatesStr = noTeammatesStr & vbLf
        noTeammatesStr = noTeammatesStr & "Waiting for other teams to finish planning."
    End If
    W_LABEL_setFormattedString d.noTeammatesLabel, noTeammatesStr, &HFFFFFFFF, font_regular
    d.noTeammatesLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.noTeammatesLabel), W_LABEL_requiredHeight(d.noTeammatesLabel)
End Sub

' PlayerNameLabel(): W_LABEL()
Public Sub PlayerNameLabel_construct(ByVal self As WIDGET)
    W_LABEL_construct self
    self.kind = K_PlayerNameLabel ' (VB) virtual calls from here on reach PlayerNameLabel's overrides
End Sub

Private Sub PlayerNameLabel_initialize(ByVal self As WIDGET, ByVal playerIdx_ As Long)
    Dim d As PlayerNameLabel
    Set d = PlayerNameLabel_(self)
    d.playerIdx = playerIdx_
    If d.playerIdx = selectedPlayer Then W_LABEL_setFont self, font_regular_bold Else W_LABEL_setFont self, font_regular
    self.setString getPlayerName(d.playerIdx)
    self.setGeometry 0, 0, W_LABEL_getMaxLineWidth(self), W_LABEL_requiredHeight(self)
    self.setTransparentToMouse True
    W_LABEL_setCanTruncate self, True
End Sub

Public Function PlayerNameLabel_make(ByVal playerIdx As Long) As WIDGET
    Dim result As WIDGET
    ' (class make_shared_enabler: public PlayerNameLabel {};)
    Set result = WIDGET_alloc(K_PlayerNameLabel)
    PlayerNameLabel_construct result
    PlayerNameLabel_initialize result, playerIdx
    Set PlayerNameLabel_make = result
End Function

Public Sub PlayerNameLabel_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    ' Update the player name (ex. once revealed in a blind_lobby game)
    self.setString getPlayerName(PlayerNameLabel_(self).playerIdx)
End Sub

' std::shared_ptr<TableRow> TeamStrategyView::newPlayerStrategyRow(uint32_t playerIdx, const std::vector<WEAPON_SUBCLASS>&
' displayedWeaponSubclasses, const std::unordered_set<WzStrategyPlanningUnitTypes>& disabledUnitTypes)
' (displayedWeaponSubclasses(0 .. displayedCount - 1); disabledUnitTypes(t): t is in the set)
Private Function TeamStrategyView_newPlayerStrategyRow(ByVal self As WIDGET, ByVal playerIdx As Long, displayedWeaponSubclasses() As Long, _
        ByVal displayedCount As Long, disabledUnitTypes() As Boolean) As WIDGET
    Dim columnWidgets As Collection, parentStrategyView_ As WIDGET, verticalRowPadding As Long, playerNameLabel As WIDGET
    Dim wrappedPlayerNameWidget As WIDGET, rowHeight As Long, i As Long, weaponSubclassWidget As WIDGET
    Dim unitType As WzStrategyPlanningUnitTypes, unitTypeWidget As WIDGET, row As WIDGET
    Set columnWidgets = New Collection

    Set parentStrategyView_ = self

    If playerIdx = selectedPlayer Then verticalRowPadding = 10 Else verticalRowPadding = 3

    ' Player Name widget
    Set playerNameLabel = PlayerNameLabel_make(playerIdx)
    Set wrappedPlayerNameWidget = Margin_wrap(verticalRowPadding, 0, verticalRowPadding, 0, playerNameLabel)  ' Margin(verticalRowPadding, 0).wrap(playerNameLabel)
    columnWidgets.Add wrappedPlayerNameWidget

    rowHeight = wrappedPlayerNameWidget.idealHeight()

    ' All of the weapon columns
    For i = 0 To displayedCount - 1
        Set weaponSubclassWidget = WzTeamStrategyWeaponButton_make(playerIdx, displayedWeaponSubclasses(i), parentStrategyView_)
        weaponSubclassWidget.setGeometry 0, 0, weaponSubclassWidget.idealWidth(), weaponSubclassWidget.idealHeight()
        rowHeight = MAX_(rowHeight, weaponSubclassWidget.height())
        columnWidgets.Add weaponSubclassWidget
    Next

    ' Add blank label (to fill the "Units:" column
    columnWidgets.Add W_LABEL_new()

    ' Unit Types
    For i = 0 To WzStrategyPlanningUnitTypes_NUM_TYPES - 1
        unitType = i
        Set unitTypeWidget = WzTeamStrategyUnitTypeButton_make(playerIdx, unitType, parentStrategyView_)
        unitTypeWidget.setGeometry 0, 0, unitTypeWidget.idealWidth(), unitTypeWidget.idealHeight()
        If disabledUnitTypes(unitType) Then
            unitTypeWidget.setState WBUT_DISABLE
        End If
        rowHeight = MAX_(rowHeight, unitTypeWidget.height())
        columnWidgets.Add unitTypeWidget
    Next

    Set row = TableRow_make(columnWidgets, rowHeight)
    TableRow_setDisabledColor row, psPalette(WZCOL_FORM_DARK)
    If playerIdx = selectedPlayer Then
        TableRow_setDrawBorder row, True, pal_RGBA(27, 42, 250, 200)
        TableRow_setHighlightsOnMouseOver row, True
    Else
        ' other player rows start off as disabled
        TableRow_setDisabled row, True
    End If
    Set TeamStrategyView_newPlayerStrategyRow = row
End Function

Private Function createColHeaderLabel(ByVal text As String, ByVal minLeftPadding As Long) As WIDGET
    Dim label As WIDGET
    Set label = W_LABEL_new()
    W_LABEL_setTextAlignment label, WLAB_ALIGNBOTTOMRIGHT
    label.setString text
    label.setGeometry 0, 0, W_LABEL_getMaxLineWidth(label), 0
    W_LABEL_setCacheNeverExpires label, True
    Set createColHeaderLabel = Margin_wrap(0, 2, WZ_TEAMSTRAT_COLUMN_HEADER_PADDING_Y, minLeftPadding, label)
End Function

Public Function TeamStrategyView_make(ByVal allowAIs As Boolean) As WIDGET
    Dim selectedPlayerTeam As Long, playersOnSameTeam() As Long, playersOnSameTeamCount As Long, player As Long
    Dim result As WIDGET, d As TeamStrategyView, columns() As TableColumn, columnsCount As Long, displayedWeaponSubclasses() As Long
    Dim displayedCount As Long, i As Long, subClass As WEAPON_SUBCLASS, weaponColumHeader As WIDGET
    Dim disabledUnitTypes(0 To WzStrategyPlanningUnitTypes_NUM_TYPES - 1) As Boolean, unitType As WzStrategyPlanningUnitTypes
    Dim typeDisabled As Boolean, unitTypeColumHeader As WIDGET, maxIdealColumnHeight As Long, minimumColumnWidths() As Long
    Dim strategyTableColumnHeaderHeight As Long, j As Long, t As Long, isSpectator As Boolean, include As Boolean
    Dim imageRef As WzTeamStrategyColumnImageRef
    If Not gameHasTeamStrategyView() Then
        Set TeamStrategyView_make = Nothing
        Exit Function
    End If

    If ULT(selectedPlayer, MAX_CONNECTED_PLAYERS) Then isSpectator = NetPlay_players_isSpectator(selectedPlayer)
    If Not ULT(selectedPlayer, MAX_PLAYERS) Or isSpectator Then
        Set TeamStrategyView_make = Nothing
        Exit Function
    End If

    ' Check if current player has teammates
    selectedPlayerTeam = checkedGetPlayerTeam(selectedPlayer)
    ReDim playersOnSameTeam(0 To MAX_PLAYERS - 1)
    For player = 0 To MIN_(game.maxPlayers, MAX_PLAYERS) - 1
        If player = selectedPlayer Then
            GoTo continue_
        End If
        If checkedGetPlayerTeam(player) <> selectedPlayerTeam Then
            GoTo continue_
        End If
        ' Only include human players (or future AIs that actually set a strategy)
        include = isHumanPlayer(player)                                                 ' is an active (connected) human player
        If Not include Then include = (NetPlay_players_difficulty(player) = AIDifficulty_HUMAN) ' was a human player (probably disconnected)
        If Not include Then include = (allowAIs And playerStrategyPlans(player).hasBeenSet) ' an AI that set a strategy (only would happen in-game, not pre-game waiting screen)
        If include Then
            playersOnSameTeam(playersOnSameTeamCount) = player
            playersOnSameTeamCount = playersOnSameTeamCount + 1
        End If
continue_:
    Next

    ' (class make_shared_enabler: public TeamStrategyView {};)
    Set result = WIDGET_alloc(K_TeamStrategyView)
    TeamStrategyView_construct result
    Set d = TeamStrategyView_(result)

    ' Create column headers for stats table
    ReDim columns(0 To WSC_NUM_WEAPON_SUBCLASSES + WzStrategyPlanningUnitTypes_NUM_TYPES + 1)
    ' "Weapons:" (right-aligned)
    columns(columnsCount) = TableColumn_(createColHeaderLabel("Weapons:", 20), TableColumn_ResizeBehavior_FIXED_WIDTH)
    columnsCount = columnsCount + 1
    ' Add weapon subclass columns
    ReDim displayedWeaponSubclasses(0 To WSC_NUM_WEAPON_SUBCLASSES - 1)
    For i = 0 To WSC_NUM_WEAPON_SUBCLASSES - 1
        subClass = i
        ' Possible Future TODO: Check if currently-loaded stats actually use this weapon subclass?
        displayedWeaponSubclasses(displayedCount) = subClass
        displayedCount = displayedCount + 1
    Next
    For i = 0 To displayedCount - 1
        subClass = displayedWeaponSubclasses(i)
        imageRef = WzTeamStrategyColumnImageRef_weapon(subClass, d.columnImagesManager)
        Set weaponColumHeader = WzTeamStrategyIconColumnHeader_make(getWeaponSubClassDisplayName(subClass, True), imageRef)
        weaponColumHeader.setTip getWeaponSubClassDisplayName(subClass, False)
        weaponColumHeader.setGeometry 0, 0, weaponColumHeader.idealWidth(), weaponColumHeader.idealHeight()
        columns(columnsCount) = TableColumn_(weaponColumHeader, TableColumn_ResizeBehavior_FIXED_WIDTH)
        columnsCount = columnsCount + 1
    Next
    ' "Units": (right-aligned)
    columns(columnsCount) = TableColumn_(createColHeaderLabel("Units:", 20), TableColumn_ResizeBehavior_FIXED_WIDTH)
    columnsCount = columnsCount + 1
    ' Add unit types columns
    For i = 0 To WzStrategyPlanningUnitTypes_NUM_TYPES - 1
        unitType = i
        typeDisabled = False
        Select Case unitType
        Case WzStrategyPlanningUnitTypes_Tanks
            typeDisabled = (ingame.flags And MPFLAGS_NO_TANKS) <> 0
        Case WzStrategyPlanningUnitTypes_VTOL
            typeDisabled = (ingame.flags And MPFLAGS_NO_VTOLS) <> 0
        Case WzStrategyPlanningUnitTypes_Cyborg
            typeDisabled = (ingame.flags And MPFLAGS_NO_CYBORGS) <> 0
        End Select
        If typeDisabled Then
            disabledUnitTypes(unitType) = True
        End If
        imageRef = WzTeamStrategyColumnImageRef_unitType(unitType, d.columnImagesManager)
        Set unitTypeColumHeader = WzTeamStrategyIconColumnHeader_make(getUnitTypeDisplayName(unitType), imageRef, typeDisabled)
        unitTypeColumHeader.setTip getUnitTypeDisplayName(unitType)
        unitTypeColumHeader.setGeometry 0, 0, unitTypeColumHeader.idealWidth(), unitTypeColumHeader.idealHeight()
        columns(columnsCount) = TableColumn_(unitTypeColumHeader, TableColumn_ResizeBehavior_FIXED_WIDTH)
        columnsCount = columnsCount + 1
    Next

    maxIdealColumnHeight = 0
    ReDim minimumColumnWidths(0 To columnsCount - 1)
    For i = 0 To columnsCount - 1
        minimumColumnWidths(i) = MAX_(columns(i).columnWidget.idealWidth(), 0)
        maxIdealColumnHeight = MAX_(maxIdealColumnHeight, columns(i).columnWidget.idealHeight())
    Next

    ' Create + attach "Strategy" scrollable table view
    strategyTableColumnHeaderHeight = MIN_(maxIdealColumnHeight, MAX_TEAMSTRAT_COLUMN_HEADERS_HEIGHT)
    Set d.table = ScrollableTableWidget_make(columns, columnsCount, strategyTableColumnHeaderHeight)
    ScrollableTableWidget_setColumnPadding d.table, Vector2i_(COL_PADDING_X, 0)
    result.attach d.table
    ScrollableTableWidget_setBackgroundColor d.table, pal_RGBA(0, 0, 0, 0)
    ScrollableTableWidget_setMinimumColumnWidths d.table, minimumColumnWidths, columnsCount
    ScrollableTableWidget_setDrawColumnLines d.table, True

    ' Add row for ourselves
    ScrollableTableWidget_addRow d.table, TeamStrategyView_newPlayerStrategyRow(result, selectedPlayer, displayedWeaponSubclasses, displayedCount, disabledUnitTypes)

    ' Add rows for all players on same team as selectedPlayer
    ' sort by player position (std::sort: an insertion sort)
    For i = 1 To playersOnSameTeamCount - 1
        t = playersOnSameTeam(i)
        j = i - 1
        Do While j >= 0
            If Not (NetPlay_players_position(t) < NetPlay_players_position(playersOnSameTeam(j))) Then Exit Do
            playersOnSameTeam(j + 1) = playersOnSameTeam(j)
            j = j - 1
        Loop
        playersOnSameTeam(j + 1) = t
    Next
    For i = 0 To playersOnSameTeamCount - 1
        player = playersOnSameTeam(i)
        If player = selectedPlayer Then
            GoTo continue2
        End If
        ScrollableTableWidget_addRow d.table, TeamStrategyView_newPlayerStrategyRow(result, player, displayedWeaponSubclasses, displayedCount, disabledUnitTypes)
continue2:
    Next
    If playersOnSameTeamCount = 0 Then
        ' Create a label below the table, since there are no other teammates to display
        Set d.noTeammatesLabel = W_LABEL_new()
        W_LABEL_setTextAlignment d.noTeammatesLabel, WLAB_ALIGNLEFT
        W_LABEL_setCacheNeverExpires d.noTeammatesLabel, True
        W_LABEL_setFontColour d.noTeammatesLabel, psPalette(WZCOL_TEXT_MEDIUM)
        TeamStrategyView_updateNoTeamLabelContents result
        result.attach d.noTeammatesLabel
        d.noTeammatesLabel.setCalcLayout WzFunc_new(FN_teamStrat_noTeammatesLabel_calcLayout)
    End If

    d.table.setCalcLayout WzFunc_new(FN_teamStrat_table_calcLayout)

    Set TeamStrategyView_make = result
End Function

Public Sub TeamStrategyView_setUpdatingPlayers(ByVal self As WIDGET, ByVal enabled As Boolean)
    Dim d As TeamStrategyView, i As Long
    Set d = TeamStrategyView_(self)
    If enabled = d.updatingPlayers Then
        Exit Sub
    End If
    d.updatingPlayers = enabled
    If d.updatingPlayers Then
        ' send any local updates to other team members
        sendStrategyPlanUpdate selectedPlayer

        ' enable all other player rows
        For i = 1 To ScrollableTableWidget_getNumRows(d.table) - 1
            ScrollableTableWidget_setRowDisabled d.table, i, False
        Next
    Else
        ' disable all other player rows
        For i = 1 To ScrollableTableWidget_getNumRows(d.table) - 1
            ScrollableTableWidget_setRowDisabled d.table, i, True
        Next
    End If
End Sub

Public Sub TeamStrategyView_setBackgroundColor(ByVal self As WIDGET, ByVal color As Long)
    TeamStrategyView_(self).backgroundColor = color
End Sub

' std::pair<std::vector<size_t>, size_t> TeamStrategyView::getMaxTableColumnDataWidths(): .first in maxColumnDataWidths
' (maxColumnDataWidthsCount of them), .second returned
Private Function TeamStrategyView_getMaxTableColumnDataWidths(ByVal self As WIDGET, maxColumnDataWidths() As Long, maxColumnDataWidthsCount As Long) As Long
    Dim d As TeamStrategyView, totalNeededColumnWidth As Long, minimumColumnWidths() As Long, minimumColumnWidthsCount As Long
    Dim colIdx As Long, numColumns As Long, maxIdealContentWidth As Long
    Set d = TeamStrategyView_(self)
    totalNeededColumnWidth = 0
    minimumColumnWidthsCount = ScrollableTableWidget_getMinimumColumnWidths(d.table, minimumColumnWidths)
    numColumns = ScrollableTableWidget_getNumColumns(d.table)
    If Not (minimumColumnWidthsCount = numColumns) Then ASSERT_FAILURE "minimumColumnWidths.size() == table->getNumColumns()", "src/hci/teamstrategy.cpp:1133", "getMaxTableColumnDataWidths", "Number of minimum column widths does not match number of columns!"
    maxColumnDataWidthsCount = 0
    If numColumns > 0 Then ReDim maxColumnDataWidths(0 To numColumns - 1)
    For colIdx = 0 To numColumns - 1
        maxIdealContentWidth = ScrollableTableWidget_getColumnMaxContentIdealWidth(d.table, colIdx)
        maxIdealContentWidth = MAX_(MAX_(maxIdealContentWidth, 0), minimumColumnWidths(colIdx))  ' (minimumColumnWidths.at(colIdx))
        If colIdx = d.playerNameColIdx Then
            ' restrict player name column to a fixed maximum width
            maxIdealContentWidth = MIN_(maxIdealContentWidth, MAX_PLAYER_NAME_COLUMN_WIDTH)
        End If
        maxColumnDataWidths(maxColumnDataWidthsCount) = maxIdealContentWidth
        maxColumnDataWidthsCount = maxColumnDataWidthsCount + 1
        totalNeededColumnWidth = totalNeededColumnWidth + maxIdealContentWidth
    Next
    TeamStrategyView_getMaxTableColumnDataWidths = totalNeededColumnWidth
End Function

Public Function TeamStrategyView_idealWidth(ByVal self As WIDGET) As Long
    Dim d As TeamStrategyView, maxColumnDataWidths() As Long, n As Long, total As Long
    Set d = TeamStrategyView_(self)
    If Not (Not d.table Is Nothing) Then ASSERT_FAILURE "table != nullptr", "src/hci/teamstrategy.cpp:1152", "idealWidth", "Table is null": TeamStrategyView_idealWidth = 0: Exit Function
    total = TeamStrategyView_getMaxTableColumnDataWidths(self, maxColumnDataWidths, n)
    TeamStrategyView_idealWidth = ScrollableTableWidget_getTableWidthNeededForTotalColumnWidth(d.table, ScrollableTableWidget_getNumColumns(d.table), total)
End Function

Public Function TeamStrategyView_idealHeight(ByVal self As WIDGET) As Long
    Dim d As TeamStrategyView, h As Long
    Set d = TeamStrategyView_(self)
    If Not d.noTeammatesLabel Is Nothing Then h = W_LABEL_requiredHeight(d.noTeammatesLabel)
    TeamStrategyView_idealHeight = d.table.idealHeight() + h
End Function

' MARK: -

' Returns true if there are *any* teams with more than one human player
Public Function gameHasTeamStrategyView(Optional ByVal allowAIs As Boolean = False) As Boolean
    Dim teamsWithIncludedPlayers As Collection, player As Long, include As Boolean, teamNumber As Long
    Set teamsWithIncludedPlayers = New Collection   ' std::unordered_set<int32_t>
    For player = 0 To MIN_(game.maxPlayers, MAX_PLAYERS) - 1
        ' Check type of player
        include = isHumanPlayer(player)                                                 ' is an active (connected) human player
        If Not include Then include = (NetPlay_players_difficulty(player) = AIDifficulty_HUMAN) ' was a human player (probably disconnected)
        If Not include Then include = (allowAIs And playerStrategyPlans(player).hasBeenSet) ' an AI that set a strategy (only would happen in-game, not pre-game waiting screen)
        If include Then
            teamNumber = checkedGetPlayerTeam(player)
            If collectionHasKey(teamsWithIncludedPlayers, "t" & teamNumber) Then
                ' more than one included player on a team
                gameHasTeamStrategyView = True
                Exit Function
            End If
            teamsWithIncludedPlayers.Add teamNumber, "t" & teamNumber
        End If
    Next
    gameHasTeamStrategyView = False
End Function

' (VB) a Collection's key test
Private Function collectionHasKey(ByVal c As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error GoTo missing
    v = c.Item(key)
    collectionHasKey = True
    Exit Function
missing:
    collectionHasKey = False
End Function

Public Function createTeamStrategyView(Optional ByVal allowAIs As Boolean = False) As WIDGET
    Set createTeamStrategyView = TeamStrategyView_make(allowAIs)
End Function

Public Function transformTeamStrategyViewMode(ByVal view As WIDGET, ByVal updatingPlayers As Boolean) As Boolean
    If view Is Nothing Then
        transformTeamStrategyViewMode = False
        Exit Function
    End If
    If Not WIDGET_isKindOf(view, K_TeamStrategyView) Then   ' std::dynamic_pointer_cast<TeamStrategyView>(view)
        transformTeamStrategyViewMode = False
        Exit Function
    End If

    TeamStrategyView_setUpdatingPlayers view, updatingPlayers
    transformTeamStrategyViewMode = True
End Function

Public Function teamStrategyViewSetBackgroundColor(ByVal view As WIDGET, ByVal color As Long) As Boolean
    If view Is Nothing Then
        teamStrategyViewSetBackgroundColor = False
        Exit Function
    End If
    If Not WIDGET_isKindOf(view, K_TeamStrategyView) Then   ' std::dynamic_pointer_cast<TeamStrategyView>(view)
        teamStrategyViewSetBackgroundColor = False
        Exit Function
    End If

    TeamStrategyView_setBackgroundColor view, color
    teamStrategyViewSetBackgroundColor = True
End Function

' void aiInformTeamOfStrategy(uint32_t aiPlayerIdx, const std::unordered_map<WEAPON_SUBCLASS, WzStrategyPlanningState>&
' weaponStrategy, const std::unordered_map<WzStrategyPlanningUnitTypes, WzStrategyPlanningState>& unitTypesStrategy)
' (the maps as keys and values: weaponKeys(i) -> weaponValues(i) for i < weaponCount, the same for the unit types)
Public Sub aiInformTeamOfStrategy(ByVal aiPlayerIdx As Long, weaponKeys() As Long, weaponValues() As Long, ByVal weaponCount As Long, _
        unitKeys() As Long, unitValues() As Long, ByVal unitCount As Long)
    Dim i As Long
    If Not ULT(aiPlayerIdx, MAX_PLAYERS) Then ASSERT_FAILURE "aiPlayerIdx < MAX_PLAYERS", "src/hci/teamstrategy.cpp:1219", "aiInformTeamOfStrategy", "Invalid aiPlayerIdx: %u", aiPlayerIdx: Exit Sub
    If Not (NetPlay_players_ai(aiPlayerIdx) >= 0) Then ASSERT_FAILURE "NetPlay.players[aiPlayerIdx].ai >= 0", "src/hci/teamstrategy.cpp:1220", "aiInformTeamOfStrategy", "Not an AI player?: %u", aiPlayerIdx: Exit Sub
    If Not (whosResponsible(aiPlayerIdx) = realSelectedPlayer) Then ASSERT_FAILURE "whosResponsible(aiPlayerIdx) == realSelectedPlayer", "src/hci/teamstrategy.cpp:1221", "aiInformTeamOfStrategy", "We are not responsible for AI player?: %u", aiPlayerIdx: Exit Sub

    WzPlayerStrategyPlan_reset aiPlayerIdx

    For i = 0 To weaponCount - 1
        playerStrategyPlans(aiPlayerIdx).weaponPlans(weaponKeys(i)).state = weaponValues(i)
    Next
    For i = 0 To unitCount - 1
        playerStrategyPlans(aiPlayerIdx).unitPlans(unitKeys(i)).state = unitValues(i)
    Next

    playerStrategyPlans(aiPlayerIdx).hasBeenSet = True

    sendStrategyPlanUpdate aiPlayerIdx
End Sub

' the lambdas
Public Function teamstrategy_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, psParent As WIDGET, y0 As Long, tableHeight As Long, pd As TeamStrategyView
    Select Case f.id
    Case FN_teamStrat_noTeammatesLabel_calcLayout   ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not psParent Is Nothing Then
            If Not WIDGET_isKindOf(psParent, K_TeamStrategyView) Then Set psParent = Nothing
        End If
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci/teamstrategy.cpp:1073", "make", "No parent": Exit Function
        ' should update the formatted string with the new maximum width
        TeamStrategyView_updateNoTeamLabelContents psParent
        y0 = psParent.height() - psWidget.height()
        psWidget.setGeometry 0, y0, psWidget.width(), psWidget.height()

    Case FN_teamStrat_table_calcLayout              ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not psParent Is Nothing Then
            If Not WIDGET_isKindOf(psParent, K_TeamStrategyView) Then Set psParent = Nothing
        End If
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci/teamstrategy.cpp:1083", "make", "No parent": Exit Function
        Set pd = TeamStrategyView_(psParent)
        tableHeight = psParent.height()
        If Not pd.noTeammatesLabel Is Nothing Then
            tableHeight = tableHeight - W_LABEL_requiredHeight(pd.noTeammatesLabel)
        End If
        psWidget.setGeometry 0, 0, psParent.width(), tableHeight
    End Select
End Function
