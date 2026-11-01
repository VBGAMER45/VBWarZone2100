Attribute VB_Name = "spectatorwidgets_c"
' src/spectatorwidgets.cpp, spectatorwidgets.h, whole: the spectators' stats overlay (the "Show Player Stats"
' button over the game and the movable stats view: each player's power and rate, power lost and won, lab use,
' kills, units, armour and weapon grades). Its widget classes' fields are the data classes named as the classes.
' WzWeaponGradesColumnManager's std::unordered_set / unordered_map members are a Boolean table and Collections
' (their order is never shown); WZ_Notification_Image(path).loadImageToTexture() is
' gfx_api_context_loadTextureFromFile, as in teamstrategy_c (whose loadImageForWeapSubclass, stats.cpp's, this
' uses). The ten label lambdas of newPlayerStatsRow are one WzFunc, FN_specStatsLabel, with the column captured.
' A std::weak_ptr captured by a lambda of a child widget is the widget's address (WIDGET_fromPtr): the child dies
' with it. Translations (_()) are not ported: the English text is used.
' (VB) widget class SpectatorStatsView : W_FORM
' (VB) widget class WzThrottledUpdateLabel : W_LABEL
' (VB) widget class WzCenteredColumnIcon : W_BUTTON
' (VB) widget class WzWeaponGradesColumnManager : WIDGET
' (VB) widget class WzWeapColumnIcon : WzCenteredColumnIcon
' (VB) widget class WeaponGradesCol : WIDGET
' (VB) widget class PlayerColorColumn : WIDGET
Option Explicit

Private Const SPEC_STATS_WINDOW_TITLE_HEIGHT As Long = 24
Private Const SPEC_STATS_WINDOW_BOTTOM_PADDING As Long = 5
Private Const MAX_PLAYER_NAME_COLUMN_WIDTH As Long = 120
Private Const PLAYER_COLOR_COL_SIZE As Long = 8
Private Const INFO_UPDATE_INTERVAL_TICKS As Long = GAME_TICKS_PER_SEC

Private Const WEAP_GRADES_COL_WIDTH As Long = 16
Private Const WEAP_GRADES_COL_IMAGE_WIDTH As Long = 16
Private Const WEAP_GRADES_COL_IMAGE_HEIGHT As Long = 16
Private Const WEAP_GRADES_COL_PADDING As Long = 6

' #define SPEC_STATS_BUTTON_X RET_X
Private Const SPEC_STATS_BUTTON_Y As Long = 20

' the lambdas (WzFunc: spectatorwidgets_c_fn)
Public Const FN_intDisplayImageHilightOnceStarted As Long = 18800 ' a WIDGET_DISPLAY (widget, xOffset, yOffset)
Public Const FN_specStatsButton_onClick As Long = 18801     ' specStatsButton->addOnClickHandler(...)
Public Const FN_specStatsViewCreate_task As Long = 18802    ' widgScheduleTask([](){ specStatsViewCreate(); })
Public Const FN_specStatsButton_hitTest As Long = 18803     ' specStatsButton->setCustomHitTest(...)
Public Const FN_specStatsViewShutdown_task As Long = 18804  ' WZ_CLEANUP_VIEW(tatsVi)
Public Const FN_weapColIcon_calcLayout As Long = 18805      ' setColImageWidgetPosition's LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_specStats_minimize_onClick As Long = 18806  ' minimizeButton->addOnClickHandler(...)
Public Const FN_specStats_toggleMinimized_task As Long = 18807 ' widgScheduleTask([psParent](){ toggleMinimized })
Public Const FN_specStats_close_onClick As Long = 18808     ' closeButton->addOnClickHandler(...)
Public Const FN_specStatsViewClose_task As Long = 18809     ' widgScheduleTask([](){ specStatsViewClose(); })
Public Const FN_specStats_closeButton_calcLayout As Long = 18810
Public Const FN_specStats_gradesWidthChange As Long = 18811 ' addOnIdealWidthChangeFunc([psWeakMain, weaponGradesManagerColIdx] ...)
Public Const FN_specStats_table_calcLayout As Long = 18812
Public Const FN_specStats_calcLayout As Long = 18813
Public Const FN_specStatsLabel As Long = 18814              ' the label lambdas: cap(0) playerIdx, cap(1) the column (1-9)

Private statsOverlay As W_SCREEN                ' static std::shared_ptr<W_SCREEN> = nullptr
Private specStatsButton As WIDGET               ' static std::shared_ptr<W_BUTTON> = nullptr
Private globalStatsForm As WIDGET               ' static std::shared_ptr<SpectatorStatsView> = nullptr

Public Sub intDisplayImageHilightOnceStarted(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    If bDisplayMultiJoiningStatus <> 0 Then
        Exit Sub
    End If
    intDisplayImageHilight psWidget, xOffset, yOffset
End Sub

' (VB) selectedPlayer < NetPlay.players.size() && NetPlay.players[selectedPlayer].isSpectator
' && realSelectedPlayer < NetPlay.players.size() && NetPlay.players[realSelectedPlayer].isSpectator
Private Function bothSelectedAreSpectators() As Boolean
    If Not ULT(selectedPlayer, NetPlay_players_size()) Then Exit Function
    If Not NetPlay_players_isSpectator(selectedPlayer) Then Exit Function
    If Not ULT(realSelectedPlayer, NetPlay_players_size()) Then Exit Function
    bothSelectedAreSpectators = NetPlay_players_isSpectator(realSelectedPlayer)
End Function

Public Function specLayerInit(Optional ByVal showButton As Boolean = True) As Boolean
    Dim x As Long, y As Long, width As Long, height As Long
    If bothSelectedAreSpectators() Then
        If statsOverlay Is Nothing Then
            Set statsOverlay = W_SCREEN_make()
            statsOverlay.psForm.hide ' hiding the root form does not stop display of children, but *does* prevent it from accepting mouse over itself - i.e. basically makes it transparent
        End If
        widgRegisterOverlayScreen statsOverlay, 65535 - 3   ' std::numeric_limits<uint16_t>::max() - 3

        ' Add the spec stats button
        If showButton And specStatsButton Is Nothing Then
            x = RET_X
            y = SPEC_STATS_BUTTON_Y
            width = iV_GetImageWidth(IntImages, IMAGE_SPECSTATS_UP)
            height = iV_GetImageHeight(IntImages, IMAGE_SPECSTATS_UP)
            If Not (width > 0 And height > 0) Then ASSERT_FAILURE "width > 0 && height > 0", "src/spectatorwidgets.cpp:116", "specLayerInit", "Possibly missing specstats button image?": specLayerInit = True: Exit Function

            Set specStatsButton = W_BUTTON_new()
            Set specStatsButton.displayFunction = WzFunc_new(FN_intDisplayImageHilightOnceStarted)
            specStatsButton.UserData = PACKDWORD_TRI(0, IMAGE_SPECSTATS_DOWN, IMAGE_SPECSTATS_UP)
            W_BUTTON_(specStatsButton).pTip = "Show Player Stats"
            W_BUTTON_addOnClickHandler specStatsButton, WzFunc_new(FN_specStatsButton_onClick)
            specStatsButton.setCustomHitTest WzFunc_new(FN_specStatsButton_hitTest)
            specStatsButton.setGeometry x, y, width, height
            statsOverlay.psForm.attach specStatsButton
        End If

        specLayerInit = True
        Exit Function
    End If

    specLayerInit = False
End Function

Private Sub specStatsViewClose()
    If Not globalStatsForm Is Nothing Then
        globalStatsForm.deleteLater
        Set globalStatsForm = Nothing
    End If

    ' Show the spec stats button
    If Not specStatsButton Is Nothing Then
        specStatsButton.show
    End If
End Sub

Public Sub specToggleOverlays()
    If Not globalStatsForm Is Nothing Then
        specStatsViewClose
    ElseIf bothSelectedAreSpectators() Then
        specStatsViewCreate
    End If
End Sub

Public Function specStatsViewShutdown() As Boolean
    If Not statsOverlay Is Nothing Then
        widgRemoveOverlayScreen statsOverlay
    End If
    specStatsViewClose
    Set specStatsButton = Nothing ' widget is owned by parent form, so will be detached / destroyed there as long as we don't keep a ref
    Set statsOverlay = Nothing
    specStatsViewShutdown = True
End Function

Public Sub specStatsViewCreate()
    Dim isSpec As Boolean
    If ULT(realSelectedPlayer, NetPlay_players_size()) Then isSpec = NetPlay_players_isSpectator(realSelectedPlayer)
    If Not isSpec Then
        If Not statsOverlay Is Nothing Then
            widgRemoveOverlayScreen statsOverlay
        End If
        Exit Sub
    End If
    If statsOverlay Is Nothing Then
        specLayerInit
        If statsOverlay Is Nothing Then
            Exit Sub
        End If
    End If
    If globalStatsForm Is Nothing Then
        Set globalStatsForm = SpectatorStatsView_make()
        If Not globalStatsForm Is Nothing Then
            statsOverlay.psForm.attach globalStatsForm
        End If
    End If
    ' Hide the spec stats button
    If Not specStatsButton Is Nothing Then
        specStatsButton.hide
    End If
End Sub

' MARK: - SpectatorStatsView

' SpectatorStatsView(): W_FORM()
Public Sub SpectatorStatsView_construct(ByVal self As WIDGET)
    W_FORM_construct self
    self.kind = K_SpectatorStatsView ' (VB) virtual calls from here on reach SpectatorStatsView's overrides
    SpectatorStatsView_(self).playerNameColIdx = 1
End Sub

Public Sub SpectatorStatsView_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), pal_RGBA(0, 0, 0, 80)
End Sub

Private Function createColHeaderLabel(ByVal text As String) As WIDGET
    Dim label As WIDGET
    Set label = W_LABEL_new()
    label.setString text
    label.setGeometry 0, 0, W_LABEL_getMaxLineWidth(label), 0
    W_LABEL_setCacheNeverExpires label, True
    Set createColHeaderLabel = label
End Function

' MARK: - WzThrottledUpdateLabel

' WzThrottledUpdateLabel(const UpdateFunc& updateFunc, uint32_t updateInterval)
Public Sub WzThrottledUpdateLabel_construct(ByVal self As WIDGET, ByVal updateFunc As WzFunc, ByVal updateInterval As Long)
    W_LABEL_construct self
    self.kind = K_WzThrottledUpdateLabel ' (VB) virtual calls from here on reach WzThrottledUpdateLabel's overrides
    Set WzThrottledUpdateLabel_(self).updateFunc = updateFunc
    WzThrottledUpdateLabel_(self).updateInterval = updateInterval
End Sub

Public Function WzThrottledUpdateLabel_make(ByVal updateFunc As WzFunc, ByVal updateInterval As Long) As WIDGET
    Dim newLabel As WIDGET
    Set newLabel = WIDGET_alloc(K_WzThrottledUpdateLabel)
    WzThrottledUpdateLabel_construct newLabel, updateFunc, updateInterval
    W_LABEL_setFont newLabel, font_regular, psPalette(WZCOL_FORM_LIGHT)
    W_LABEL_setCanTruncate newLabel, False
    newLabel.setTransparentToClicks True
    If Not updateFunc Is Nothing Then
        WzFunc_call updateFunc, newLabel
    End If
    newLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(newLabel), iV_GetTextLineSize(font_regular)
    Set WzThrottledUpdateLabel_make = newLabel
End Function

Public Sub WzThrottledUpdateLabel_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzThrottledUpdateLabel
    W_LABEL_run self, psContext
    Set d = WzThrottledUpdateLabel_(self)
    If Not ULT(gameTime - d.lastGameTimeUpdated, d.updateInterval) Then   ' gameTime - lastGameTimeUpdated >= updateInterval
        If Not d.updateFunc Is Nothing Then WzFunc_call d.updateFunc, self
        d.lastGameTimeUpdated = gameTime
    End If
End Sub

Private Function loadImageForArmorCol(ByVal thermal As Boolean) As gfx_api_texture
    Dim imagePath As String
    imagePath = "images/intfac/armor_kinetic.png"
    If thermal Then
        imagePath = "images/intfac/armor_thermal.png"
    End If
    ' WZ_Notification_Image img(imagePath); return img.loadImageToTexture();
    Set loadImageForArmorCol = gfx_api_context_loadTextureFromFile(imagePath, texture_type_user_interface)
End Function

' MARK: - WzCenteredColumnIcon

' WzCenteredColumnIcon(gfx_api::texture* pImageTexture, int displayWidth, int displayHeight)
' (~WzCenteredColumnIcon deletes the texture: released with the data object)
Public Sub WzCenteredColumnIcon_construct(ByVal self As WIDGET, ByVal pImageTexture As gfx_api_texture, ByVal displayWidth As Long, ByVal displayHeight As Long)
    Dim d As WzCenteredColumnIcon
    W_BUTTON_construct self
    self.kind = K_WzCenteredColumnIcon ' (VB) virtual calls from here on reach WzCenteredColumnIcon's overrides
    Set d = WzCenteredColumnIcon_(self)
    Set d.pImageTexture = pImageTexture
    d.displayWidth = displayWidth
    d.displayHeight = displayHeight
End Sub

Public Function WzCenteredColumnIcon_new(ByVal pImageTexture As gfx_api_texture, ByVal displayWidth As Long, ByVal displayHeight As Long) As WIDGET
    Set WzCenteredColumnIcon_new = WIDGET_alloc(K_WzCenteredColumnIcon)
    WzCenteredColumnIcon_construct WzCenteredColumnIcon_new, pImageTexture, displayWidth, displayHeight
End Function

Public Sub WzCenteredColumnIcon_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzCenteredColumnIcon, x0 As Long, y0 As Long, imageX0 As Long, imageY0 As Long
    Set d = WzCenteredColumnIcon_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    If Not d.pImageTexture Is Nothing Then
        imageX0 = x0 + (self.width() - d.displayWidth) \ 2
        imageY0 = y0 + (self.height() - d.displayHeight) \ 2
        iV_DrawImageAnisotropic d.pImageTexture, imageX0, imageY0, 0!, 0!, d.displayWidth, d.displayHeight, 0!, pal_RGBA(255, 255, 255, 200)
    End If
End Sub

' MARK: - WzWeaponGradesColumnManager

' WzWeaponGradesColumnManager(): WIDGET()
Public Sub WzWeaponGradesColumnManager_construct(ByVal self As WIDGET)
    Dim d As WzWeaponGradesColumnManager
    WIDGET_construct self
    self.kind = K_WzWeaponGradesColumnManager ' (VB) virtual calls from here on reach WzWeaponGradesColumnManager's overrides
    Set d = WzWeaponGradesColumnManager_(self)
    Set d.cachedNumberTexts = New Collection
    Set d.columnIcons = New Collection
    Set d.onIdealWidthChangeFuncs = New Collection
End Sub

Public Function WzWeaponGradesColumnManager_make() As WIDGET
    Dim result As WIDGET, d As WzWeaponGradesColumnManager
    ' (class make_shared_enabler)
    Set result = WIDGET_alloc(K_WzWeaponGradesColumnManager)
    WzWeaponGradesColumnManager_construct result
    Set d = WzWeaponGradesColumnManager_(result)

    WzWeaponGradesColumnManager_updateData result
    d.idealHeightForEachRow = iV_GetTextLineSize(font_regular)
    result.setGeometry 0, 0, 64, MAX_(d.idealHeightForEachRow, WEAP_GRADES_COL_WIDTH)

    Set WzWeaponGradesColumnManager_make = result
End Function

Public Sub WzWeaponGradesColumnManager_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzWeaponGradesColumnManager
    Set d = WzWeaponGradesColumnManager_(self)
    If Not ULT(gameTime - d.lastGameTimeUpdated, d.updateInterval) Then   ' gameTime - lastGameTimeUpdated >= updateInterval
        WzWeaponGradesColumnManager_updateData self
        d.lastGameTimeUpdated = gameTime
        If Not NetPlay_players_isSpectator(realSelectedPlayer) Then
            widgScheduleTask WzFunc_new(FN_specStatsViewShutdown_task)  ' WZ_CLEANUP_VIEW(tatsVi)
        End If
    End If
End Sub

' (display: a no-op, as WIDGET's)

Public Sub WzWeaponGradesColumnManager_geometryChanged(ByVal self As WIDGET)
    Dim d As WzWeaponGradesColumnManager, icon As WIDGET, i As Long
    Set d = WzWeaponGradesColumnManager_(self)
    If d.columnIcons Is Nothing Then Exit Sub   ' (VB) inside WIDGET's constructor
    For i = 1 To d.columnIcons.Count
        Set icon = d.columnIcons.Item(i)
        icon.callCalcLayout
    Next
End Sub

Public Function WzWeaponGradesColumnManager_idealWidth(ByVal self As WIDGET) As Long
    Dim numColumns As Long
    numColumns = WzWeaponGradesColumnManager_(self).visibleColumnCount
    WzWeaponGradesColumnManager_idealWidth = (numColumns * WEAP_GRADES_COL_WIDTH) + (MAX_(numColumns - 1, 0) * WEAP_GRADES_COL_PADDING)
End Function

Public Function WzWeaponGradesColumnManager_idealHeight(ByVal self As WIDGET) As Long
    WzWeaponGradesColumnManager_idealHeight = MAX_(WzWeaponGradesColumnManager_(self).idealHeightForEachRow, WEAP_GRADES_COL_WIDTH)
End Function

Public Sub WzWeaponGradesColumnManager_addOnIdealWidthChangeFunc(ByVal self As WIDGET, ByVal handler As WzFunc)
    If handler Is Nothing Then
        Exit Sub
    End If
    WzWeaponGradesColumnManager_(self).onIdealWidthChangeFuncs.Add handler
End Sub

Private Function WzWeaponGradesColumnManager_getColumnLeftPositionFromIndex(ByVal colIndex As Long) As Long
    WzWeaponGradesColumnManager_getColumnLeftPositionFromIndex = colIndex * (WEAP_GRADES_COL_WIDTH + WEAP_GRADES_COL_PADDING)
End Function

Private Function WzWeaponGradesColumnManager_getWzCachedTextForNumber(ByVal self As WIDGET, ByVal number As Long) As WzCachedText
    Dim d As WzWeaponGradesColumnManager, newCachedText As WzCachedText
    Set d = WzWeaponGradesColumnManager_(self)
    ' should be self-limiting because there are only a limited number of grades for any weapon
    On Error Resume Next
    Set newCachedText = d.cachedNumberTexts.Item("n" & number)
    On Error GoTo 0
    If Not newCachedText Is Nothing Then
        Set WzWeaponGradesColumnManager_getWzCachedTextForNumber = newCachedText
        Exit Function
    End If
    Set newCachedText = WzCachedText_new_text(CStr(number), font_regular)
    d.cachedNumberTexts.Add newCachedText, "n" & number
    Set WzWeaponGradesColumnManager_getWzCachedTextForNumber = newCachedText
End Function

Private Sub WzWeaponGradesColumnManager_buildWeaponSubclassInfo(ByVal self As WIDGET)
    Dim d As WzWeaponGradesColumnManager, i As Long, playerIdx As Long, gradeForPlayer As Long
    Set d = WzWeaponGradesColumnManager_(self)
    For i = 0 To WSC_NUM_WEAPON_SUBCLASSES - 1
        For playerIdx = 0 To game.maxPlayers - 1
            gradeForPlayer = getNumWeaponImpactClassUpgrades(playerIdx, i)
            If gradeForPlayer = d.maxNumWeaponGrade(i) Then
                d.hasMaxGrade(i, playerIdx) = True
            ElseIf gradeForPlayer > d.maxNumWeaponGrade(i) Then
                d.clearPlayersWithMaxGrade i
                d.maxNumWeaponGrade(i) = gradeForPlayer
                d.hasMaxGrade(i, playerIdx) = True
            End If
        Next
    Next
End Sub

' (bool hiddenColumn(WEAPON_SUBCLASS subClass): always false)

Private Sub WzWeaponGradesColumnManager_setColImageWidgetPosition(ByVal self As WIDGET, ByVal subClass As WEAPON_SUBCLASS, ByVal columnIdx As Long)
    Dim d As WzWeaponGradesColumnManager, icon As WIDGET
    Set d = WzWeaponGradesColumnManager_(self)
    On Error Resume Next
    Set icon = d.columnIcons.Item("s" & subClass)
    On Error GoTo 0
    If icon Is Nothing Then
        Set icon = WzWeapColumnIcon_make(subClass)
        d.columnIcons.Add icon, "s" & subClass
        self.attach icon
        icon.setCalcLayout WzFunc_new(FN_weapColIcon_calcLayout)
    End If
    icon.setGeometry WzWeaponGradesColumnManager_getColumnLeftPositionFromIndex(columnIdx), icon.y(), icon.width(), icon.height()
End Sub

Private Sub WzWeaponGradesColumnManager_calculateColumnPositions(ByVal self As WIDGET)
    Dim d As WzWeaponGradesColumnManager, i As Long, icon As WIDGET
    Set d = WzWeaponGradesColumnManager_(self)
    d.visibleColumnCount = 0
    For i = 0 To WSC_NUM_WEAPON_SUBCLASSES - 1
        If d.maxNumWeaponGrade(i) = 0 Then   ' hiddenColumn(subClass) || info.maxNumWeaponGrade == 0
            ' hide child widget
            Set icon = Nothing
            On Error Resume Next
            Set icon = d.columnIcons.Item("s" & i)
            On Error GoTo 0
            If Not icon Is Nothing Then
                icon.hide
            End If
            GoTo continue_
        End If
        d.visibleColumn(d.visibleColumnCount) = i
        d.visibleColumnCount = d.visibleColumnCount + 1
        ' reposition child widget for icon
        WzWeaponGradesColumnManager_setColImageWidgetPosition self, i, d.visibleColumnCount - 1
continue_:
    Next
End Sub

Private Sub WzWeaponGradesColumnManager_updateData(ByVal self As WIDGET)
    Dim d As WzWeaponGradesColumnManager, idealWidthBefore As Long, cachedText As WzCachedText, i As Long
    Set d = WzWeaponGradesColumnManager_(self)
    idealWidthBefore = self.idealWidth()
    WzWeaponGradesColumnManager_buildWeaponSubclassInfo self
    WzWeaponGradesColumnManager_calculateColumnPositions self
    For i = 1 To d.cachedNumberTexts.Count
        Set cachedText = d.cachedNumberTexts.Item(i)
        cachedText.tick
    Next
    If self.idealWidth() <> idealWidthBefore Then
        For i = 1 To d.onIdealWidthChangeFuncs.Count
            WzFunc_call d.onIdealWidthChangeFuncs.Item(i), self
        Next
    End If
End Sub

' MARK: - WzWeapColumnIcon

Public Sub WzWeapColumnIcon_construct(ByVal self As WIDGET, ByVal pImageTexture As gfx_api_texture, ByVal displayWidth As Long, ByVal displayHeight As Long)
    WzCenteredColumnIcon_construct self, pImageTexture, displayWidth, displayHeight
    self.kind = K_WzWeapColumnIcon ' (VB) virtual calls from here on reach WzWeapColumnIcon's overrides
End Sub

Public Function WzWeapColumnIcon_make(ByVal subClass As WEAPON_SUBCLASS) As WIDGET
    Dim pImageTexture As gfx_api_texture, result As WIDGET
    ' Load the required image for display
    Set pImageTexture = loadImageForWeapSubclass(subClass)

    ' (class make_shared_enabler)
    Set result = WIDGET_alloc(K_WzWeapColumnIcon)
    WzWeapColumnIcon_construct result, pImageTexture, WEAP_GRADES_COL_IMAGE_WIDTH, WEAP_GRADES_COL_IMAGE_HEIGHT
    WzWeapColumnIcon_(result).subClass = subClass
    Set WzWeapColumnIcon_make = result
End Function

Public Function WzWeapColumnIcon_getTip(ByVal self As WIDGET) As String
    Dim tip As String
    tip = "Weapon Grade"
    tip = tip & vbLf
    tip = tip & getWeaponSubClass_name(WzWeapColumnIcon_(self).subClass)
    WzWeapColumnIcon_getTip = tip
End Function

' MARK: - WeaponGradesCol

' WeaponGradesCol(const std::shared_ptr<WzWeaponGradesColumnManager>& manager, uint32_t playerIdx)
Public Sub WeaponGradesCol_construct(ByVal self As WIDGET, ByVal manager As WIDGET, ByVal playerIdx As Long)
    WIDGET_construct self
    self.kind = K_WeaponGradesCol ' (VB) virtual calls from here on reach WeaponGradesCol's overrides
    Set WeaponGradesCol_(self).manager = manager
    WeaponGradesCol_(self).playerIdx = playerIdx
End Sub

Public Sub WeaponGradesCol_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WeaponGradesCol, md As WzWeaponGradesColumnManager, x0 As Long, i As Long, subClass As Long
    Dim behindOtherPlayers As Boolean, text As WzCachedText, posX0 As Long, textBoundingBoxOffset_y As Long
    Dim fy As Single, textColor As Long
    Set d = WeaponGradesCol_(self)
    If Not (Not d.manager Is Nothing) Then ASSERT_FAILURE "manager != nullptr", "src/spectatorwidgets.cpp:589", "display", "null manager": Exit Sub
    Set md = WzWeaponGradesColumnManager_(d.manager)
    x0 = self.x() + xOffset

    For i = 0 To md.visibleColumnCount - 1
        subClass = md.visibleColumn(i)
        behindOtherPlayers = Not md.hasMaxGrade(subClass, d.playerIdx)     ' playersWithMaxGrade.count(playerIdx) == 0
        Set text = WzWeaponGradesColumnManager_getWzCachedTextForNumber(d.manager, getNumWeaponImpactClassUpgrades(d.playerIdx, subClass))
        ' colRect = manager->getColumnRectFromIndex(i): (left, 0, WEAP_GRADES_COL_WIDTH, WEAP_GRADES_COL_IMAGE_HEIGHT)
        posX0 = x0 + WzWeaponGradesColumnManager_getColumnLeftPositionFromIndex(i) + (WEAP_GRADES_COL_WIDTH - text.deref().width()) \ 2
        textBoundingBoxOffset_y = yOffset + self.y() + (self.height() - text.deref().lineSize()) \ 2
        fy = CSng(textBoundingBoxOffset_y) - CSng(text.deref().aboveBase())
        If Not behindOtherPlayers Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_TEXT_MEDIUM)
        text.deref().render posX0, fy, textColor
    Next
End Sub

Public Function WeaponGradesCol_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WeaponGradesCol
    Set d = WeaponGradesCol_(self)
    If Not (Not d.manager Is Nothing) Then ASSERT_FAILURE "manager != nullptr", "src/spectatorwidgets.cpp:609", "idealWidth", "null manager": WeaponGradesCol_idealWidth = 0: Exit Function
    WeaponGradesCol_idealWidth = d.manager.idealWidth()
End Function

' MARK: - PlayerColorColumn

' PlayerColorColumn(uint32_t playerIdx)
Public Sub PlayerColorColumn_construct(ByVal self As WIDGET, ByVal playerIdx As Long)
    WIDGET_construct self
    self.kind = K_PlayerColorColumn ' (VB) virtual calls from here on reach PlayerColorColumn's overrides
    PlayerColorColumn_(self).playerIdx = playerIdx
End Sub

Public Sub PlayerColorColumn_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, posX0 As Long, posY0 As Long, playerColor As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    posX0 = x0 + 1
    posY0 = y0

    playerColor = pal_GetTeamColour(getPlayerColour(PlayerColorColumn_(self).playerIdx))
    playerColor = (playerColor And &HFFFFFF) Or &HDC000000   ' playerColor.byte.a = 220
    pie_UniTransBoxFill posX0, posY0, posX0 + self.width(), posY0 + (self.height() - 1), playerColor
End Sub

Public Function PlayerColorColumn_idealHeight(ByVal self As WIDGET) As Long
    PlayerColorColumn_idealHeight = PLAYER_COLOR_COL_SIZE + 1
End Function

Public Function PlayerColorColumn_idealWidth(ByVal self As WIDGET) As Long
    PlayerColorColumn_idealWidth = PLAYER_COLOR_COL_SIZE
End Function

' MARK: - SpectatorStatsView::newPlayerStatsRow

Private Function SpectatorStatsView_newPlayerStatsRow(ByVal self As WIDGET, ByVal playerIdx As Long, Optional ByVal rowHeight As Long = 0) As WIDGET
    Dim columnWidgets As Collection, playerColWidget As WIDGET, column As Long, weapGradesCol As WIDGET, manager As WIDGET
    Set columnWidgets = New Collection

    ' Player color widget
    Set playerColWidget = WIDGET_alloc(K_PlayerColorColumn)
    PlayerColorColumn_construct playerColWidget, playerIdx
    playerColWidget.setGeometry 0, 0, PLAYER_COLOR_COL_SIZE, PLAYER_COLOR_COL_SIZE + 1
    columnWidgets.Add playerColWidget

    ' Player Name, Power info, PowerLost, PowerWon, Lab Use, Kills, Units (as in F5), Kinetic Armor (T/C),
    ' Thermal Armor (T/C): the label lambdas (FN_specStatsLabel's columns 1 to 9)
    For column = 1 To 9
        columnWidgets.Add WzThrottledUpdateLabel_make(WzFunc_new(FN_specStatsLabel, playerIdx, column), INFO_UPDATE_INTERVAL_TICKS)
    Next
    ' Weapon Grades
    Set manager = SpectatorStatsView_(self).weaponGradesManager
    Set weapGradesCol = WIDGET_alloc(K_WeaponGradesCol)
    WeaponGradesCol_construct weapGradesCol, manager, playerIdx
    weapGradesCol.setGeometry 0, 0, manager.idealWidth(), WzWeaponGradesColumnManager_(manager).idealHeightForEachRow
    columnWidgets.Add weapGradesCol

    Set SpectatorStatsView_newPlayerStatsRow = TableRow_make(columnWidgets, rowHeight)
End Function

' a label lambda of newPlayerStatsRow ([playerIdx](W_LABEL& label){ ... }), by its column
Private Sub specStatsLabelUpdate(ByVal label As WIDGET, ByVal playerIdx As Long, ByVal column As Long)
    Dim sts As PLAYERSTATS_
    Select Case column
    Case 1  ' Player Name widget
        label.setString getPlayerName(playerIdx)
    Case 2  ' Power info
        label.setString CStr(getPower(playerIdx)) & "/" & getApproxPowerGeneratedPerSecForDisplay(playerIdx)
    Case 3  ' PowerLost
        label.setString CStr(getMultiStats(playerIdx).recentPowerLost)
    Case 4  ' PowerWon
        label.setString CStr(getMultiStats(playerIdx).recentPowerWon)
    Case 5  ' Lab Use
        sts = getMultiStats(playerIdx)
        If sts.recentResearchPotential = 0 Then
            label.setString " - "
        Else
            label.setString va("%.2f%%", (CSng(sts.recentResearchPerformance) / CSng(sts.recentResearchPotential)) * 100)
        End If
    Case 6  ' Kills
        label.setString CStr(getMultiStats(playerIdx).recentKills)
    Case 7  ' Units (as in F5)
        label.setString CStr(getNumDroids(playerIdx) + getNumTransporterDroids(playerIdx))
    Case 8  ' Kinetic Armor (T/C)
        label.setString CStr(getNumBodyClassArmourUpgrades(playerIdx, BodyClass_Tank)) & "/" & CStr(getNumBodyClassArmourUpgrades(playerIdx, BodyClass_Cyborg))
    Case 9  ' Thermal Armor (T/C)
        label.setString CStr(getNumBodyClassThermalArmourUpgrades(playerIdx, BodyClass_Tank)) & "/" & CStr(getNumBodyClassThermalArmourUpgrades(playerIdx, BodyClass_Cyborg))
    End Select
    ' ADJUST_LABEL_COLOR_FOR_PLAYERS()
    If (Not NetPlay_players_allocated(playerIdx) And NetPlay_players_ai(playerIdx) < 0) Or NetPlay_players_isSpectator(playerIdx) Then
        W_LABEL_setFontColour label, psPalette(WZCOL_TEXT_MEDIUM)
    End If
End Sub

' MARK: - SpectatorStatsView::make

Public Function SpectatorStatsView_make() As WIDGET
    Dim result As WIDGET, d As SpectatorStatsView, newFormWidth As Long, minimizeButton As WIDGET, sButInit As W_BUTINIT
    Dim columns(0 To 10) As TableColumn, armorLabel As WIDGET, thermalArmorLabel As WIDGET, weaponGradesManagerColIdx As Long
    Dim minimumColumnWidths(0 To 10) As Long, i As Long, playerIndexes(0 To MAX_PLAYERS - 1) As Long, playerCount As Long
    Dim player As Long, j As Long, t As Long, label As WIDGET
    ' (class make_shared_enabler)
    Set result = WIDGET_alloc(K_SpectatorStatsView)
    SpectatorStatsView_construct result
    Set d = SpectatorStatsView_(result)

    ' make minimizable
    W_FORM_enableMinimizing result, "Player Stats", psPalette(WZCOL_FORM_TEXT)

    ' Main "parent" form
    newFormWidth = FRONTEND_TOPFORM_WIDEW
    result.setGeometry (pie_GetVideoBufferWidth() - newFormWidth) \ 2, 10, newFormWidth, pie_GetVideoBufferHeight() - 50
    W_FORM_(result).userMovable = True

    ' Add the "minimize" button
    Set minimizeButton = makeFormTransparentCornerButton(ChrW$(&H21B8), 10, pal_RGBA(0, 0, 0, 60)) ' U+21B8
    result.attach minimizeButton
    W_BUTTON_addOnClickHandler minimizeButton, WzFunc_new(FN_specStats_minimize_onClick)

    ' Add the "close" button
    W_BUTINIT_construct sButInit
    sButInit.id = IDOBJ_CLOSE
    sButInit.pTip = "Close"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_CLOSEHILIGHT, IMAGE_CLOSE)
    Set d.closeButton = W_BUTTON_new_init(sButInit)
    result.attach d.closeButton
    W_BUTTON_addOnClickHandler d.closeButton, WzFunc_new(FN_specStats_close_onClick)
    d.closeButton.setCalcLayout WzFunc_new(FN_specStats_closeButton_calcLayout)

    If Not bothSelectedAreSpectators() Then Exit Function   ' (return nullptr)

    ' Create column headers for stats table
    columns(0) = TableColumn_(createColHeaderLabel(""), TableColumn_ResizeBehavior_FIXED_WIDTH)
    columns(1) = TableColumn_(createColHeaderLabel("Player"), TableColumn_ResizeBehavior_RESIZABLE)
    Set label = createColHeaderLabel("Power/Rate")
    label.setTip "Current Power / Power Per Second"
    columns(2) = TableColumn_(label, TableColumn_ResizeBehavior_RESIZABLE)
    columns(3) = TableColumn_(createColHeaderLabel("PwrLost"), TableColumn_ResizeBehavior_RESIZABLE)
    columns(4) = TableColumn_(createColHeaderLabel("PwrWon"), TableColumn_ResizeBehavior_RESIZABLE)
    columns(5) = TableColumn_(createColHeaderLabel("Lab Use"), TableColumn_ResizeBehavior_RESIZABLE)
    columns(6) = TableColumn_(createColHeaderLabel("Kills"), TableColumn_ResizeBehavior_RESIZABLE)
    columns(7) = TableColumn_(createColHeaderLabel("Units"), TableColumn_ResizeBehavior_RESIZABLE)
    ' createArmorColWidget(thermal): the image loaded for display
    Set armorLabel = WzCenteredColumnIcon_new(loadImageForArmorCol(False), WEAP_GRADES_COL_IMAGE_WIDTH, WEAP_GRADES_COL_IMAGE_HEIGHT)
    armorLabel.setTip "Kinetic Armor" & vbLf & "(Tanks / Cyborgs)"
    Set thermalArmorLabel = WzCenteredColumnIcon_new(loadImageForArmorCol(True), WEAP_GRADES_COL_IMAGE_WIDTH, WEAP_GRADES_COL_IMAGE_HEIGHT)
    thermalArmorLabel.setTip "Thermal Armor" & vbLf & "(Tanks / Cyborgs)"
    Set d.weaponGradesManager = WzWeaponGradesColumnManager_make()
    columns(8) = TableColumn_(armorLabel, TableColumn_ResizeBehavior_FIXED_WIDTH)
    columns(9) = TableColumn_(thermalArmorLabel, TableColumn_ResizeBehavior_FIXED_WIDTH)
    columns(10) = TableColumn_(d.weaponGradesManager, TableColumn_ResizeBehavior_FIXED_WIDTH)
    weaponGradesManagerColIdx = 10      ' columns.size() - 1
    For i = 0 To 10
        minimumColumnWidths(i) = MAX_(columns(i).columnWidget.idealWidth(), 0)
    Next
    minimumColumnWidths(0) = PLAYER_COLOR_COL_SIZE
    WzWeaponGradesColumnManager_addOnIdealWidthChangeFunc d.weaponGradesManager, WzFunc_new(FN_specStats_gradesWidthChange, ObjPtr(result), weaponGradesManagerColIdx)

    ' Create + attach "Triggers" scrollable table view
    Set d.table = ScrollableTableWidget_make(columns, 11)
    result.attach d.table
    ScrollableTableWidget_setBackgroundColor d.table, pal_RGBA(0, 0, 0, 25)
    ScrollableTableWidget_setMinimumColumnWidths d.table, minimumColumnWidths, 11

    ' Add rows for all players
    For player = 0 To MIN_(game.maxPlayers, MAX_PLAYERS) - 1
        If isHumanPlayer(player) Or NetPlay_players_difficulty(player) = AIDifficulty_HUMAN Or NetPlay_players_ai(player) >= 0 Then
            ' is an active (connected) human player, was a human player (probably disconnected), is an AI bot
            playerIndexes(playerCount) = player
            playerCount = playerCount + 1
        End If
        If player = selectedPlayer And Not NetPlay_players_isSpectator(player) Then
            Set SpectatorStatsView_make = Nothing
            Exit Function
        End If
    Next
    ' sort by player position (std::sort: an insertion sort)
    For i = 1 To playerCount - 1
        t = playerIndexes(i)
        j = i - 1
        Do While j >= 0
            If Not (NetPlay_players_position(t) < NetPlay_players_position(playerIndexes(j))) Then Exit Do
            playerIndexes(j + 1) = playerIndexes(j)
            j = j - 1
        Loop
        playerIndexes(j + 1) = t
    Next
    For i = 0 To playerCount - 1
        player = playerIndexes(i)
        If player = selectedPlayer And Not NetPlay_players_isSpectator(player) Then
            Set SpectatorStatsView_make = Nothing
            Exit Function
        End If
        ScrollableTableWidget_addRow d.table, SpectatorStatsView_newPlayerStatsRow(result, player)
    Next

    d.table.setCalcLayout WzFunc_new(FN_specStats_table_calcLayout)

    ' layout calculations for main form
    result.setCalcLayout WzFunc_new(FN_specStats_calcLayout)

    Set SpectatorStatsView_make = result
End Function

' std::pair<std::vector<size_t>, size_t> SpectatorStatsView::getMaxTableColumnDataWidths(): .first in maxColumnDataWidths
' (maxColumnDataWidthsCount of them), .second returned
Private Function SpectatorStatsView_getMaxTableColumnDataWidths(ByVal self As WIDGET, maxColumnDataWidths() As Long, maxColumnDataWidthsCount As Long) As Long
    Dim d As SpectatorStatsView, totalNeededColumnWidth As Long, minimumColumnWidths() As Long, minimumColumnWidthsCount As Long
    Dim colIdx As Long, numColumns As Long, maxIdealContentWidth As Long
    Set d = SpectatorStatsView_(self)
    totalNeededColumnWidth = 0
    minimumColumnWidthsCount = ScrollableTableWidget_getMinimumColumnWidths(d.table, minimumColumnWidths)
    numColumns = ScrollableTableWidget_getNumColumns(d.table)
    If Not (minimumColumnWidthsCount = numColumns) Then ASSERT_FAILURE "minimumColumnWidths.size() == table->getNumColumns()", "src/spectatorwidgets.cpp:897", "getMaxTableColumnDataWidths", "Number of minimum column widths does not match number of columns!"
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
    SpectatorStatsView_getMaxTableColumnDataWidths = totalNeededColumnWidth
End Function

Private Sub SpectatorStatsView_resizeTableColumnWidths(ByVal self As WIDGET)
    Dim maxColumnDataWidths() As Long, n As Long
    SpectatorStatsView_getMaxTableColumnDataWidths self, maxColumnDataWidths, n
    ScrollableTableWidget_changeColumnWidths SpectatorStatsView_(self).table, maxColumnDataWidths, n
End Sub

Public Sub SpectatorStatsView_geometryChanged(ByVal self As WIDGET)
    Dim d As SpectatorStatsView
    Set d = SpectatorStatsView_(self)
    If Not d.closeButton Is Nothing Then
        d.closeButton.callCalcLayout
    End If
    If Not d.table Is Nothing Then
        d.table.callCalcLayout
    End If
End Sub

Public Function SpectatorStatsView_idealWidth(ByVal self As WIDGET) As Long
    Dim d As SpectatorStatsView, maxColumnDataWidths() As Long, n As Long, total As Long
    Set d = SpectatorStatsView_(self)
    If Not (Not d.table Is Nothing) Then ASSERT_FAILURE "table != nullptr", "src/spectatorwidgets.cpp:934", "idealWidth", "Table is null": SpectatorStatsView_idealWidth = 0: Exit Function
    total = SpectatorStatsView_getMaxTableColumnDataWidths(self, maxColumnDataWidths, n)
    SpectatorStatsView_idealWidth = ScrollableTableWidget_getTableWidthNeededForTotalColumnWidth(d.table, ScrollableTableWidget_getNumColumns(d.table), total)
End Function

Public Function SpectatorStatsView_idealHeight(ByVal self As WIDGET) As Long
    Dim d As SpectatorStatsView, h As Long
    Set d = SpectatorStatsView_(self)
    If Not d.table Is Nothing Then h = d.table.idealHeight()
    SpectatorStatsView_idealHeight = SPEC_STATS_WINDOW_TITLE_HEIGHT + h + SPEC_STATS_WINDOW_BOTTOM_PADDING
End Function

' (VB) std::dynamic_pointer_cast<SpectatorStatsView>(w): w when it is one, else Nothing
Private Function asSpectatorStatsView(ByVal w As WIDGET) As WIDGET
    If w Is Nothing Then Exit Function
    If WIDGET_isKindOf(w, K_SpectatorStatsView) Then Set asSpectatorStatsView = w
End Function

' the lambdas
Public Function spectatorwidgets_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, psParent As WIDGET, pd As SpectatorStatsView, oldWidth As Long, y0 As Long, x0 As Long
    Dim newWidth As Long, newHeight As Long, manager As WIDGET
    Select Case f.id
    Case FN_intDisplayImageHilightOnceStarted
        intDisplayImageHilightOnceStarted args(0), args(1), args(2)

    Case FN_specStatsButton_onClick                 ' [](W_BUTTON& button)
        widgScheduleTask WzFunc_new(FN_specStatsViewCreate_task)

    Case FN_specStatsViewCreate_task
        specStatsViewCreate

    Case FN_specStatsButton_hitTest                 ' [](const WIDGET *psWidget, int x, int y) -> bool
        If bDisplayMultiJoiningStatus <> 0 Then
            ' effectively: make this unclickable until the game actually starts
            spectatorwidgets_c_fn = False
        Else
            spectatorwidgets_c_fn = True ' revert back to default hit-testing behavior
        End If

    Case FN_specStatsViewShutdown_task              ' WZ_CLEANUP_VIEW(tatsVi): specStatsViewShutdown()
        specStatsViewShutdown

    Case FN_weapColIcon_calcLayout                  ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not psParent Is Nothing Then
            If Not WIDGET_isKindOf(psParent, K_WzWeaponGradesColumnManager) Then Set psParent = Nothing
        End If
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/spectatorwidgets.cpp:469", "setColImageWidgetPosition", "null parent": Exit Function
        psWidget.setGeometry psWidget.x(), 0, WEAP_GRADES_COL_IMAGE_WIDTH, psParent.height()

    Case FN_specStats_minimize_onClick              ' [](W_BUTTON& button)
        Set psWidget = args(0)
        Set psParent = asSpectatorStatsView(psWidget.parent())
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/spectatorwidgets.cpp:746", "make", "No parent": Exit Function
        widgScheduleTask WzFunc_new(FN_specStats_toggleMinimized_task, psParent)

    Case FN_specStats_toggleMinimized_task          ' [psParent]()
        W_FORM_toggleMinimized f.cap(0)

    Case FN_specStats_close_onClick                 ' [](W_BUTTON& button)
        Set psWidget = args(0)
        Set psParent = asSpectatorStatsView(psWidget.parent())
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/spectatorwidgets.cpp:762", "make", "No parent": Exit Function
        widgScheduleTask WzFunc_new(FN_specStatsViewClose_task)

    Case FN_specStatsViewClose_task
        specStatsViewClose

    Case FN_specStats_closeButton_calcLayout        ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = asSpectatorStatsView(psWidget.parent())
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/spectatorwidgets.cpp:769", "make", "No parent": Exit Function
        psWidget.setGeometry psParent.width() - CLOSE_WIDTH, 0, CLOSE_WIDTH, CLOSE_HEIGHT

    Case FN_specStats_gradesWidthChange             ' [psWeakMain, weaponGradesManagerColIdx](WzWeaponGradesColumnManager& manager)
        Set manager = args(0)
        Set psParent = WIDGET_fromPtr(f.cap(0))     ' psWeakMain.lock()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psStrongMain != nullptr", "src/spectatorwidgets.cpp:820", "make", "Null main view": Exit Function
        ScrollableTableWidget_setMinimumColumnWidth SpectatorStatsView_(psParent).table, f.cap(1), manager.idealWidth()
        psParent.callCalcLayout
        SpectatorStatsView_resizeTableColumnWidths psParent

    Case FN_specStats_table_calcLayout              ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = asSpectatorStatsView(psWidget.parent())
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/spectatorwidgets.cpp:862", "make", "No parent": Exit Function
        Set pd = SpectatorStatsView_(psParent)
        oldWidth = psWidget.width()
        y0 = SPEC_STATS_WINDOW_TITLE_HEIGHT
        psWidget.setGeometry 0, y0, psParent.width(), MIN_(psParent.height() - (y0 + SPEC_STATS_WINDOW_BOTTOM_PADDING), pd.table.idealHeight())

        If oldWidth <> psWidget.width() Then
            SpectatorStatsView_resizeTableColumnWidths psParent
        End If

    Case FN_specStats_calcLayout                    ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = asSpectatorStatsView(psWidget)   ' psStrongSelf
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psStrongSelf != nullptr", "src/spectatorwidgets.cpp:876", "make", "Null??": Exit Function
        ' update the width
        newWidth = MIN_(pie_GetVideoBufferWidth() - 50, psParent.idealWidth())
        ' update the height
        newHeight = MIN_(pie_GetVideoBufferHeight() - 50, psParent.idealHeight())
        ' update the x if necessary to ensure the entire form is visible
        x0 = MIN_(psWidget.x(), pie_GetVideoBufferWidth() - newWidth)
        ' update the y if necessary to ensure the entire form is visible
        y0 = MIN_(psWidget.y(), pie_GetVideoBufferHeight() - newHeight)

        psWidget.setGeometry x0, y0, newWidth, newHeight

    Case FN_specStatsLabel                          ' [playerIdx](W_LABEL& label)
        specStatsLabelUpdate args(0), f.cap(0), f.cap(1)
    End Select
End Function
