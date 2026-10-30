Attribute VB_Name = "multilobbyoptions_c"
' src/titleui/widgets/multilobbyoptions.cpp, multilobbyoptions.h: the lobby's
' options form (WzMultiLobbyOptionsImpl: the status bar, room name, map,
' match options, limits), its option buttons and the password pop-over. The
' widget object model is docs/CONVENTIONS.md's "Widgets": the classes' fields
' are their data classes (same names), the lambdas and display functions
' WzFuncs run by multilobbyoptions_c_fn. The weak pointers: a
' WzMultiLobbyOptionsImpl or WzLobbyPasswordOptionPopover is its serial
' (widgweak_vb: the pop-over's close handler can run after the form is gone),
' the others addresses (ObjPtr) of widgets that own the lambda. std::array
' limitIcons is limitIcons(i) and limitIcons_size(). addMultiOptButtonT<T>'s T
' is a parameter (MULTIOPT_T_*), and multivote's overloads by the option's
' enum are suffixed with the enum (getMultiOptionPrefValue_ScavType, ...).
' (VB) widget class WzMultiLobbyOptionsWidgetBase : WIDGET
' (VB) virtual WzMultiLobbyOptionsWidgetBase.refreshData
' (VB) widget class WzMultiLobbyStatusBar : WIDGET
' (VB) widget class LobbyMultiOptionsSectionWidget : MultibuttonWidget
' (VB) widget class WzLobbyMultiOptionButton : W_BUTTON
' (VB) widget class LobbyMultiOptionsChoiceWidget : MultibuttonWidget
' (VB) widget class WzLobbyPasswordOptionPopover : WIDGET
' (VB) widget class WzMultiLobbyOptionsImpl : WzMultiLobbyOptionsWidgetBase
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' struct LimitIcon
Public Type LimitIcon
    stat As String                      ' char const *: vbNullString is nullptr
    desc As String                      ' char const *
    icon As Long
End Type

' WzLobbyMultiOptionButton::MultiOptionButtonMode
Public Enum MultiOptionButtonMode
    MultiOptionButtonMode_HOST_CONFIGURATION
    MultiOptionButtonMode_CLIENT
End Enum

Private Const MultiOptionsRowHeight As Long = 24

' (VB) addMultiOptButtonT<T>'s T
Private Const MULTIOPT_T_ScavType As Long = 0
Private Const MULTIOPT_T_AllianceType As Long = 1
Private Const MULTIOPT_T_PowerSetting As Long = 2
Private Const MULTIOPT_T_CampType As Long = 3
Private Const MULTIOPT_T_TechLevel As Long = 4

' the lambdas and display functions
Public Const FN_displayPasswordEditBox As Long = 6300
Public Const FN_MLO_PW_RETURN As Long = 6301            ' passwordEditBox->setOnReturnHandler([weakPopover](W_EDITBOX& widg) {...})
Public Const FN_MLO_PW_ESCAPE As Long = 6302            ' passwordEditBox->setOnEscapeHandler([weakPopover](W_EDITBOX& widg) {...})
Public Const FN_MLO_PW_CHECKBOX_CLICK As Long = 6303    ' requiredPasswordCheckbox->addOnClickHandler([weakPopover](W_BUTTON&) {...})
Public Const FN_displayWrappedMainRoomItem As Long = 6304
Public Const FN_displayWrappedOptionSectionItem As Long = 6305
Public Const FN_displayInlineEditBox As Long = 6306
Public Const FN_gameNameEditHandlerFunc As Long = 6307
Public Const FN_MLO_OPTION_IS_HOST As Long = 6308       ' auto optionIsHost = [](MultibuttonWidget&, int newValue) -> bool {...}
Public Const FN_MLO_GAMENAME_ESCAPE As Long = 6309      ' gameNameEditBox->setOnEscapeHandler([](W_EDITBOX& widg) {...})
Public Const FN_MLO_ROOMNAME_EDIT_CLICK As Long = 6310  ' roomNameEditButton->addOnClickHandler([psWeakGameNameEditBox](W_BUTTON&) {...})
Public Const FN_MLO_ROOMNAME_EDIT_TASK As Long = 6311
Public Const FN_MLO_PASSWORD_BUTTON_CLICK As Long = 6312 ' roomPasswordOptionsButton->addOnClickHandler([psWeakMultiLobbyOptions](W_BUTTON&) {...})
Public Const FN_MLO_PASSWORD_BUTTON_TASK As Long = 6313
Public Const FN_MLO_MAPVIEW_CLICK As Long = 6314
Public Const FN_MLO_MAPVIEW_TASK As Long = 6315
Public Const FN_MLO_MAPCHANGE_CLICK As Long = 6316      ' mapChangeButton->addOnClickHandler([psWeakTitleUI](W_BUTTON&) {...})
Public Const FN_MLO_MATCHOPTIONS_CLICK As Long = 6317   ' matchOptionsConfigureButton->addOnClickHandler([psWeakMultiLobbyOptions](W_BUTTON&) {...})
Public Const FN_MLO_MATCHOPTIONS_TASK As Long = 6318
Public Const FN_MLO_SCAV_CANCHOOSE As Long = 6319
Public Const FN_MLO_SCAV_ONCHOOSE As Long = 6320
Public Const FN_MLO_ALLIANCE_CANCHOOSE As Long = 6321
Public Const FN_MLO_ALLIANCE_ONCHOOSE As Long = 6322
Public Const FN_MLO_POWER_CANCHOOSE As Long = 6323
Public Const FN_MLO_POWER_ONCHOOSE As Long = 6324
Public Const FN_MLO_BASE_CANCHOOSE As Long = 6325
Public Const FN_MLO_BASE_ONCHOOSE As Long = 6326
Public Const FN_MLO_TECH_ONCHOOSE As Long = 6327
Public Const FN_MLO_LIMITSINFO_CLICK As Long = 6328
Public Const FN_MLO_LIMITSINFO_TASK As Long = 6329
Public Const FN_MLO_LIMITSINFO_PRINTLINE As Long = 6330 ' printStructureLimitsInfo's [](const std::string& line) {...}
Public Const FN_MLO_LIMITSOPTIONS_CLICK As Long = 6331
Public Const FN_MLO_LIMITSOPTIONS_TASK As Long = 6332
Public Const FN_MLO_OPTBUTTON_UPDATE As Long = 6333     ' addMultiOptButtonT's updateDataFunc = [pref, weakSelf]() {...}: cap(0) T, cap(1) pref, cap(2) weakSelf
Public Const FN_MLO_OPTBUTTON_CLICK As Long = 6334      ' button->addOnClickHandler([pref](W_BUTTON& b) {...}): cap(0) T, cap(1) pref
Public Const FN_MLO_PWOVERLAY_CLOSE As Long = 6335      ' displayPasswordOverlayForm's close handler [weakPasswordPopover, weakMultiLobbyOptions]()
Public Const FN_MLO_PWOVERLAY_FOCUS_TASK As Long = 6336 ' widgScheduleTask([weakPasswordPopover]() {...})
Public Const FN_MLO_MENUITEM_CLICK As Long = 6337       ' addClickableMenuItem's [onClick](WzClickableOptionsChoiceWidget&, WIDGET_KEY key)
Public Const FN_MLO_MENU_RANDOMIZE As Long = 6338       ' "Randomize Options": [](WIDGET_KEY)
Public Const FN_MLO_MENU_STARALL As Long = 6339         ' "Star All": [](WIDGET_KEY)
Public Const FN_MLO_MENU_CLEARALL As Long = 6340        ' "Clear All": [](WIDGET_KEY)
Public Const FN_MLO_MENU_SHOWPREFS As Long = 6341       ' "Show / Hide Player Preferences": [weakMultiLobbyOptions](WIDGET_KEY): cap(1) the value set
Public Const FN_MLO_MATCHMENU_CLOSE As Long = 6342      ' displayMatchOptionsMenu's close handler [weakMultiLobbyOptions]()

' (VB) PIELIGHT's byte.a, read and set
Private Function pielightA(ByVal c As Long) As Long
    pielightA = SHRU(c, 24) And &HFF&
End Function

Private Function pielightWithA(ByVal c As Long, ByVal a As Long) As Long
    pielightWithA = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' MARK: -

' std::array<LimitIcon, 6> limitIcons
Public Function limitIcons(ByVal i As Long) As LimitIcon
    Select Case i
    Case 0: limitIcons.stat = "A0LightFactory": limitIcons.desc = "Tanks disabled!!": limitIcons.icon = IMAGE_NO_TANK
    Case 1: limitIcons.stat = "A0CyborgFactory": limitIcons.desc = "Cyborgs disabled.": limitIcons.icon = IMAGE_NO_CYBORG
    Case 2: limitIcons.stat = "A0VTolFactory1": limitIcons.desc = "VTOLs disabled.": limitIcons.icon = IMAGE_NO_VTOL
    Case 3: limitIcons.stat = "A0Sat-linkCentre": limitIcons.desc = "Satellite Uplink disabled.": limitIcons.icon = IMAGE_NO_UPLINK
    Case 4: limitIcons.stat = "A0LasSatCommand": limitIcons.desc = "Laser Satellite disabled.": limitIcons.icon = IMAGE_NO_LASSAT
    Case 5: limitIcons.stat = vbNullString: limitIcons.desc = "Structure Limits Enforced.": limitIcons.icon = IMAGE_DARK_LOCKED
    End Select
End Function

' limitIcons.size()
Public Function limitIcons_size() As Long
    limitIcons_size = 6
End Function

' MARK: - WzMultiLobbyOptionsWidgetBase

Public Sub WzMultiLobbyOptionsWidgetBase_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_WzMultiLobbyOptionsWidgetBase ' (VB) virtual calls from here on reach WzMultiLobbyOptionsWidgetBase's overrides
End Sub

' virtual void refreshData() = 0; (VB) a pure virtual: every concrete class overrides it, so this is never called
Public Sub WzMultiLobbyOptionsWidgetBase_refreshData(ByVal self As WIDGET)
End Sub

' MARK: - WzMultiLobbyStatusBar

Public Sub WzMultiLobbyStatusBar_construct(ByVal self As WIDGET)
    Dim d As WzMultiLobbyStatusBar
    WIDGET_construct self
    self.kind = K_WzMultiLobbyStatusBar ' (VB) virtual calls from here on reach WzMultiLobbyStatusBar's overrides
    Set d = WzMultiLobbyStatusBar_(self)
    d.paddingX = 16
    d.paddingY = 5
    d.isChallenge = False
End Sub

Public Function WzMultiLobbyStatusBar_make(ByVal isChallenge As Boolean) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzMultiLobbyStatusBar)
    WzMultiLobbyStatusBar_construct widget
    WzMultiLobbyStatusBar_initialize widget, isChallenge
    Set WzMultiLobbyStatusBar_make = widget
End Function

Private Sub WzMultiLobbyStatusBar_initialize(ByVal self As WIDGET, ByVal isChallengeVal As Boolean)
    Dim d As WzMultiLobbyStatusBar
    Set d = WzMultiLobbyStatusBar_(self)
    d.isChallenge = isChallengeVal

    Set d.leftStatus = W_LABEL_new()
    W_LABEL_setFont d.leftStatus, font_small, psPalette(WZCOL_TEXT_MEDIUM)
    self.attach d.leftStatus

    Set d.rightStatus = W_LABEL_new()
    W_LABEL_setFont d.rightStatus, font_small, psPalette(WZCOL_TEXT_MEDIUM)
    W_LABEL_setTextAlignment d.rightStatus, WLAB_ALIGNRIGHT
    self.attach d.rightStatus

    WzMultiLobbyStatusBar_refreshData self
End Sub

Public Sub WzMultiLobbyStatusBar_refreshData(ByVal self As WIDGET)
    Dim d As WzMultiLobbyStatusBar, leftStatusString As String, rightStatusString As String
    Set d = WzMultiLobbyStatusBar_(self)
    If ingame.side = InGameSide_HOST_OR_SINGLEPLAYER Then
        If Not NetPlay_bComms Then ' bComms == true when a true multiplayer game (false if skirmish or challenge)
            If d.isChallenge Then leftStatusString = "Challenge:" Else leftStatusString = "Skirmish:"
            rightStatusString = "Local (Offline)"
        Else
            leftStatusString = "Hosting:"
            If NetPlay_GamePassworded() Then
                rightStatusString = "Password-Protected (Private)"
            Else
                rightStatusString = "Open (Public)"
            End If
        End If
    Else
        leftStatusString = "Connected to Host:"
        rightStatusString = NetPlay_players_name(NetPlay_hostPlayer)
    End If

    d.leftStatus.setString leftStatusString
    d.rightStatus.setString rightStatusString
End Sub

Public Sub WzMultiLobbyStatusBar_geometryChanged(ByVal self As WIDGET)
    Dim d As WzMultiLobbyStatusBar, usableWidth As Long, widgetX0 As Long, widgetHeight As Long, widgetWidth As Long
    Set d = WzMultiLobbyStatusBar_(self)
    usableWidth = self.width() - (d.paddingX * 2)
    widgetX0 = d.paddingX
    widgetHeight = self.height()
    widgetWidth = (usableWidth \ 2) - (d.paddingX \ 2)

    d.leftStatus.setGeometry widgetX0, 0, widgetWidth, widgetHeight

    widgetX0 = self.width() - d.paddingX - widgetWidth
    d.rightStatus.setGeometry widgetX0, 0, widgetWidth, widgetHeight
End Sub

Public Function WzMultiLobbyStatusBar_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WzMultiLobbyStatusBar
    Set d = WzMultiLobbyStatusBar_(self)
    WzMultiLobbyStatusBar_idealWidth = d.paddingX + W_LABEL_getMaxLineWidth(d.leftStatus) + d.paddingX + W_LABEL_getMaxLineWidth(d.rightStatus) + d.paddingX
End Function

Public Function WzMultiLobbyStatusBar_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzMultiLobbyStatusBar, a As Long, b As Long
    Set d = WzMultiLobbyStatusBar_(self)
    a = W_LABEL_requiredHeight(d.leftStatus)
    b = W_LABEL_requiredHeight(d.rightStatus)
    If b > a Then a = b
    WzMultiLobbyStatusBar_idealHeight = (d.paddingY * 2) + a
End Function

Public Sub WzMultiLobbyStatusBar_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    ' Draw dark background
    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), pal_RGBA(0, 0, 0, 255)
End Sub

' MARK: - LobbyMultiOptionsSectionWidget

Public Sub LobbyMultiOptionsSectionWidget_construct(ByVal self As WIDGET)
    MultibuttonWidget_construct self, -1
    self.kind = K_LobbyMultiOptionsSectionWidget ' (VB) virtual calls from here on reach LobbyMultiOptionsSectionWidget's overrides
End Sub

Public Function LobbyMultiOptionsSectionWidget_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_LobbyMultiOptionsSectionWidget)
    LobbyMultiOptionsSectionWidget_construct w
    Set LobbyMultiOptionsSectionWidget_new = w
End Function

Public Sub LobbyMultiOptionsSectionWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' no-op
End Sub

Public Function LobbyMultiOptionsSectionWidget_idealWidth(ByVal self As WIDGET) As Long
    LobbyMultiOptionsSectionWidget_idealWidth = MultibuttonWidget_idealWidth(self)
End Function

Public Function LobbyMultiOptionsSectionWidget_idealHeight(ByVal self As WIDGET) As Long
    LobbyMultiOptionsSectionWidget_idealHeight = MultiOptionsRowHeight
End Function

Public Sub LobbyMultiOptionsSectionWidget_stateChanged(ByVal self As WIDGET)
    ' no-op
End Sub

' MARK: - LobbyMultiOptionsChoiceWidget

Public Sub WzLobbyMultiOptionButton_construct(ByVal self As WIDGET)
    Dim d As WzLobbyMultiOptionButton
    W_BUTTON_construct self
    self.kind = K_WzLobbyMultiOptionButton ' (VB) virtual calls from here on reach WzLobbyMultiOptionButton's overrides
    Set d = WzLobbyMultiOptionButton_(self)
    d.mode = MultiOptionButtonMode_HOST_CONFIGURATION
    d.voteTotal = 0
    d.showSelected = False
    d.showTotalVotes = False
    d.localVoteIndicator = False
    d.showLocalVoteIndicatorOnHighlight = False
    d.voteIndicatorColor = pal_RGBA(13, 255, 0, 255)
    d.voteIndicatorDimensions = 12
    Set d.voteTotalText = New WzCachedText
End Sub

Public Function WzLobbyMultiOptionButton_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_WzLobbyMultiOptionButton)
    WzLobbyMultiOptionButton_construct w
    Set WzLobbyMultiOptionButton_new = w
End Function

Public Sub WzLobbyMultiOptionButton_setMode(ByVal self As WIDGET, ByVal newMode As MultiOptionButtonMode)
    WzLobbyMultiOptionButton_(self).mode = newMode
End Sub

Public Sub WzLobbyMultiOptionButton_setVoteTotal(ByVal self As WIDGET, ByVal votes As Long)
    Dim d As WzLobbyMultiOptionButton
    Set d = WzLobbyMultiOptionButton_(self)
    d.voteTotal = votes
    If d.voteTotal > 0 Then
        d.voteTotalText.setText CStr(votes), font_small
    Else
        d.voteTotalText.setText "", font_small
    End If
End Sub

Public Sub WzLobbyMultiOptionButton_setShowTotalVotes(ByVal self As WIDGET, ByVal show As Boolean)
    WzLobbyMultiOptionButton_(self).showTotalVotes = show
End Sub

Public Sub WzLobbyMultiOptionButton_setLocalVoteIndicator(ByVal self As WIDGET, ByVal thumbsUp As Boolean)
    WzLobbyMultiOptionButton_(self).localVoteIndicator = thumbsUp
End Sub

Public Sub WzLobbyMultiOptionButton_setShowLocalVoteIndicatorOnHighlight(ByVal self As WIDGET, ByVal val_ As Boolean)
    WzLobbyMultiOptionButton_(self).showLocalVoteIndicatorOnHighlight = val_
End Sub

Public Sub WzLobbyMultiOptionButton_setShowSelected(ByVal self As WIDGET, ByVal val_ As Boolean)
    WzLobbyMultiOptionButton_(self).showSelected = val_
End Sub

Public Sub WzLobbyMultiOptionButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzLobbyMultiOptionButton, x0 As Long, y0 As Long, w As Long, st As Long, imgs As W_BUTTON_Images
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, overlay_color As Long, a As Long
    Dim textWidth As Long, textX0 As Long, textY0 As Long, textX1 As Long
    Dim textBorderX0 As Long, textBorderX1 As Long, textBorderY0 As Long, textBorderY1 As Long
    Dim voteColor As Long, imgPosX0 As Long, imgPosY0 As Long
    Set d = WzLobbyMultiOptionButton_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    st = W_BUTTON_(self).state
    imgs = W_BUTTON_(self).images

    isDown = ((st And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0 Or d.showSelected)
    isDisabled = (st And WBUT_DISABLE) <> 0 Or d.mode = MultiOptionButtonMode_CLIENT
    isHighlight = (st And WBUT_HIGHLIGHT) <> 0

    ' Display the button.
    If Not AtlasImage_isNull(imgs.normal) Then
        iV_DrawImageImage imgs.normal.images, imgs.normal.id, x0, y0
        If isDown And Not AtlasImage_isNull(imgs.down) Then
            iV_DrawImageImage imgs.down.images, imgs.down.id, x0, y0
        End If
        If isDisabled And Not AtlasImage_isNull(imgs.disabled) Then
            iV_DrawImageImage imgs.disabled.images, imgs.disabled.id, x0, y0
        End If
        If isHighlight And Not AtlasImage_isNull(imgs.highlighted) Then
            iV_DrawImageImage imgs.highlighted.images, imgs.highlighted.id, x0, y0
        End If
    End If

    If isDisabled And Not AtlasImage_isNull(imgs.normal) And AtlasImage_isNull(imgs.disabled) Then
        ' disabled, render something over it!
        overlay_color = psPalette(WZCOL_TRANSPARENT_BOX)
        a = pielightA(overlay_color)
        If a > 120 Then a = 120
        If a < 70 Then a = 70
        overlay_color = pielightWithA(overlay_color, a)
        If d.showSelected Then
            overlay_color = pielightWithA(overlay_color, pielightA(overlay_color) \ 3)
        End If
        pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), overlay_color
    End If

    If d.showTotalVotes And d.voteTotal <> 0 Then
        ' Draw vote totals on right side
        textWidth = d.voteTotalText.deref().width()
        textX0 = x0 + w - textWidth - 3
        textY0 = y0 + (self.height() - d.voteTotalText.deref().lineSize()) - d.voteTotalText.deref().aboveBase() - 1
        textX1 = textX0 + textWidth

        textBorderX0 = textX0 - 1
        If textBorderX0 < x0 Then textBorderX0 = x0
        textBorderX1 = textX1 + 1
        If textBorderX1 > x0 + w Then textBorderX1 = x0 + w
        textBorderY0 = textY0 + d.voteTotalText.deref().aboveBase()
        textBorderY1 = textBorderY0 + d.voteTotalText.deref().lineSize()
        pie_UniTransBoxFill textBorderX0, textBorderY0, textBorderX1, textBorderY1, psPalette(WZCOL_TRANSPARENT_BOX)
        d.voteTotalText.deref().render textX0 - 1, textY0 + 1, psPalette(WZCOL_BLACK)
        d.voteTotalText.deref().render textX0, textY0, d.voteIndicatorColor
    Else
        d.voteTotalText.tick
    End If

    If d.localVoteIndicator Or (isHighlight And d.showLocalVoteIndicatorOnHighlight) Then
        ' Draw local vote indicator on left side
        voteColor = d.voteIndicatorColor
        If Not d.localVoteIndicator Then
            voteColor = pielightWithA(voteColor, 120)
        End If
        imgPosX0 = x0 + 2
        imgPosY0 = y0 + (self.height() - d.voteIndicatorDimensions) - 2
        pie_UniTransBoxFill imgPosX0, imgPosY0, imgPosX0 + d.voteIndicatorDimensions, imgPosY0 + d.voteIndicatorDimensions, psPalette(WZCOL_TRANSPARENT_BOX)
        iV_DrawImageFileAnisotropicTint FrontImages, IMAGE_STAR_FILL, imgPosX0 - 1, imgPosY0 + 1, d.voteIndicatorDimensions, d.voteIndicatorDimensions, psPalette(WZCOL_BLACK)
        iV_DrawImageFileAnisotropicTint FrontImages, IMAGE_STAR_FILL, imgPosX0, imgPosY0, d.voteIndicatorDimensions, d.voteIndicatorDimensions, voteColor
    End If
End Sub

' LobbyMultiOptionsChoiceWidget(bool editMode)
Public Sub LobbyMultiOptionsChoiceWidget_construct(ByVal self As WIDGET, ByVal editMode As Boolean)
    MultibuttonWidget_construct self, -1
    self.kind = K_LobbyMultiOptionsChoiceWidget ' (VB) virtual calls from here on reach LobbyMultiOptionsChoiceWidget's overrides
    LobbyMultiOptionsChoiceWidget_(self).editMode = editMode
    MultibuttonWidget_(self).lockCurrent = editMode
End Sub

Public Function LobbyMultiOptionsChoiceWidget_new(ByVal editMode As Boolean) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_LobbyMultiOptionsChoiceWidget)
    LobbyMultiOptionsChoiceWidget_construct w, editMode
    Set LobbyMultiOptionsChoiceWidget_new = w
End Function

Public Sub LobbyMultiOptionsChoiceWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' no-op
End Sub

Public Function LobbyMultiOptionsChoiceWidget_isInEditMode(ByVal self As WIDGET) As Boolean
    LobbyMultiOptionsChoiceWidget_isInEditMode = LobbyMultiOptionsChoiceWidget_(self).editMode
End Function

Public Sub LobbyMultiOptionsChoiceWidget_stateChanged(ByVal self As WIDGET)
    Dim m As MultibuttonWidget, i As Long, button As WIDGET, isSelectedItem As Boolean, st As Long
    Set m = MultibuttonWidget_(self)
    For i = 1 To m.buttons.Count
        Set button = m.buttons(i)
        isSelectedItem = m.buttonValues(i) = m.currentValue_
        If isSelectedItem And m.lockCurrent Then
            st = WBUT_LOCK
        ElseIf m.disabled Then
            st = WBUT_DISABLE
        Else
            st = 0
        End If
        button.setState st

        If WIDGET_isKindOf(button, K_WzLobbyMultiOptionButton) Then
            WzLobbyMultiOptionButton_setShowSelected button, isSelectedItem
        End If
    Next
End Sub

' MARK: - WzLobbyPasswordOptionPopover

Public Sub WzLobbyPasswordOptionPopover_construct(ByVal self As WIDGET)
    Dim d As WzLobbyPasswordOptionPopover
    WIDGET_construct self
    self.kind = K_WzLobbyPasswordOptionPopover ' (VB) virtual calls from here on reach WzLobbyPasswordOptionPopover's overrides
    Set d = WzLobbyPasswordOptionPopover_(self)
    d.paddingX = 15
    d.paddingY = 15
    d.innerPaddingY = 10
    d.editBoxHeight = 24
    d.weakSerial = WeakWidget_register(self)
End Sub

Public Sub WzLobbyPasswordOptionPopover_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzLobbyPasswordOptionPopover_(self).weakSerial
End Sub

' static std::shared_ptr<WzLobbyPasswordOptionPopover> make(const std::weak_ptr<WzMultiLobbyOptionsWidgetBase>& parentOptions, const WzString& defaultPassword = "")
' (parentOptions: the WzMultiLobbyOptionsImpl's serial)
Public Function WzLobbyPasswordOptionPopover_make(ByVal parentOptions As Long, Optional ByVal defaultPassword As String = "") As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzLobbyPasswordOptionPopover)
    WzLobbyPasswordOptionPopover_construct widget
    WzLobbyPasswordOptionPopover_initialize widget, parentOptions, defaultPassword
    Set WzLobbyPasswordOptionPopover_make = widget
End Function

Private Sub displayPasswordEditBox(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, colour As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    drawBlueBox x, y, psWidget.width(), psWidget.height()

    If (W_EDITBOX_(psWidget).state And WEDBS_DISABLE) <> 0 Then ' disabled
        ' (PIELIGHT colour: byte.r = FILLRED, byte.b = FILLBLUE, byte.g = FILLGREEN, byte.a = FILLTRANS)
        colour = pal_RGBA(FILLRED, FILLGREEN, FILLBLUE, FILLTRANS)
        pie_UniTransBoxFill x, y, x + psWidget.width() + psWidget.height() - 1, y + psWidget.height(), colour
    End If
End Sub

Private Sub WzLobbyPasswordOptionPopover_initialize(ByVal self As WIDGET, ByVal parentOptions As Long, ByVal defaultPassword As String)
    Dim d As WzLobbyPasswordOptionPopover, weakPopover As Long, s As String
    Set d = WzLobbyPasswordOptionPopover_(self)
    d.weakParentMultiOptions = parentOptions

    weakPopover = d.weakSerial

    Set d.passwordEditBox = W_EDITBOX_new()
    Set W_EDITBOX_(d.passwordEditBox).pBoxDisplay = WzFunc_new(FN_displayPasswordEditBox)
    W_EDITBOX_setPlaceholder d.passwordEditBox, "Set Password"
    W_EDITBOX_setPlaceholderTextColor d.passwordEditBox, True, psPalette(WZCOL_TEXT_DARK)
    If Len(NetPlay_gamePassword()) > 0 Then s = NetPlay_gamePassword() Else s = defaultPassword
    d.passwordEditBox.setString s
    W_EDITBOX_setMaxStringSize d.passwordEditBox, password_string_size - 1
    W_EDITBOX_setOnReturnHandler d.passwordEditBox, WzFunc_new(FN_MLO_PW_RETURN, weakPopover)
    W_EDITBOX_setOnEscapeHandler d.passwordEditBox, WzFunc_new(FN_MLO_PW_ESCAPE, weakPopover)
    self.attach d.passwordEditBox

    Set d.requiredPasswordCheckbox = WzAdvCheckbox_make("Require Password to Join", "")
    WzAdvCheckbox_setOuterHorizontalPadding d.requiredPasswordCheckbox, 0
    WzAdvCheckbox_setInnerHorizontalPadding d.requiredPasswordCheckbox, 10
    WzAdvCheckbox_setOuterVerticalPadding d.requiredPasswordCheckbox, 0
    WzAdvCheckbox_setIsChecked d.requiredPasswordCheckbox, NETGameIsLocked()
    If Len(d.passwordEditBox.getString()) = 0 Then
        d.requiredPasswordCheckbox.setState WBUT_DISABLE
    Else
        d.requiredPasswordCheckbox.setState 0
    End If
    W_BUTTON_addOnClickHandler d.requiredPasswordCheckbox, WzFunc_new(FN_MLO_PW_CHECKBOX_CLICK, weakPopover)
    self.attach d.requiredPasswordCheckbox
End Sub

Private Sub WzLobbyPasswordOptionPopover_setGamePassword(ByVal self As WIDGET, ByVal newPass As String)
    Dim d As WzLobbyPasswordOptionPopover, buf As String, strongParentMultiOptions As WIDGET
    Set d = WzLobbyPasswordOptionPopover_(self)
    ' (char buf[255])
    NETsetGamePassword newPass
    buf = Left$("*** password [" & NetPlay_gamePassword() & "] is now required! ***", 254)
    displayRoomNotifyMessage buf
    WzAdvCheckbox_setIsChecked d.requiredPasswordCheckbox, True
    d.requiredPasswordCheckbox.setState 0
    ' Update parent form about state change
    Set strongParentMultiOptions = WeakWidget_lock(d.weakParentMultiOptions)
    If Not strongParentMultiOptions Is Nothing Then
        vt_WzMultiLobbyOptionsWidgetBase_refreshData strongParentMultiOptions
    End If
End Sub

Private Sub WzLobbyPasswordOptionPopover_resetGamePassword(ByVal self As WIDGET)
    Dim d As WzLobbyPasswordOptionPopover, buf As String, strongParentMultiOptions As WIDGET
    Set d = WzLobbyPasswordOptionPopover_(self)
    ' (char buf[255])
    NETresetGamePassword
    buf = "*** password is NOT required! ***"
    displayRoomNotifyMessage buf
    WzAdvCheckbox_setIsChecked d.requiredPasswordCheckbox, False
    If Len(d.passwordEditBox.getString()) = 0 Then
        d.requiredPasswordCheckbox.setState WBUT_DISABLE
    Else
        d.requiredPasswordCheckbox.setState 0
    End If
    ' Update parent form about state change
    Set strongParentMultiOptions = WeakWidget_lock(d.weakParentMultiOptions)
    If Not strongParentMultiOptions Is Nothing Then
        vt_WzMultiLobbyOptionsWidgetBase_refreshData strongParentMultiOptions
    End If
End Sub

Public Sub WzLobbyPasswordOptionPopover_givePasswordFieldFocus(ByVal self As WIDGET)
    Dim d As WzLobbyPasswordOptionPopover, context As W_CONTEXT
    Set d = WzLobbyPasswordOptionPopover_(self)
    ' give the edit box the focus
    ' (W_CONTEXT context = W_CONTEXT::ZeroContext(): a new W_CONTEXT is zeroed)
    context.mx = d.passwordEditBox.screenPosX() + d.passwordEditBox.width()
    context.my = d.passwordEditBox.screenPosY()
    W_EDITBOX_simulateClick d.passwordEditBox, context
End Sub

Public Function WzLobbyPasswordOptionPopover_getCurrentPasswordFieldValue(ByVal self As WIDGET) As String
    WzLobbyPasswordOptionPopover_getCurrentPasswordFieldValue = WzLobbyPasswordOptionPopover_(self).passwordEditBox.getString()
End Function

Public Function WzLobbyPasswordOptionPopover_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WzLobbyPasswordOptionPopover, a As Long
    Set d = WzLobbyPasswordOptionPopover_(self)
    a = d.requiredPasswordCheckbox.idealWidth()
    If a < 150 Then a = 150
    WzLobbyPasswordOptionPopover_idealWidth = d.paddingX + a + d.paddingX
End Function

Public Function WzLobbyPasswordOptionPopover_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzLobbyPasswordOptionPopover
    Set d = WzLobbyPasswordOptionPopover_(self)
    WzLobbyPasswordOptionPopover_idealHeight = d.paddingY + d.editBoxHeight + d.innerPaddingY + d.requiredPasswordCheckbox.idealHeight() + d.paddingY
End Function

Public Sub WzLobbyPasswordOptionPopover_geometryChanged(ByVal self As WIDGET)
    Dim d As WzLobbyPasswordOptionPopover, widgetX0 As Long, widgetY0 As Long, usableWidth As Long, usableHeight As Long
    Set d = WzLobbyPasswordOptionPopover_(self)
    widgetX0 = d.paddingX
    widgetY0 = d.paddingY
    usableWidth = self.width() - (d.paddingX * 2)
    usableHeight = self.height() - (d.paddingY * 2)

    If usableWidth <= 0 Or usableHeight <= 0 Then
        Exit Sub
    End If

    d.passwordEditBox.setGeometry widgetX0, widgetY0, usableWidth, d.editBoxHeight

    widgetY0 = d.passwordEditBox.y() + d.passwordEditBox.height() + d.innerPaddingY
    d.requiredPasswordCheckbox.setGeometry widgetX0, widgetY0, usableWidth, d.requiredPasswordCheckbox.idealHeight()
End Sub

Public Sub WzLobbyPasswordOptionPopover_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    drawBlueBox x0, y0, self.width(), self.height()
End Sub

' MARK: - Helpers

Public Function formatMapName(ByVal name_ As String) As String
    Dim withoutTechlevel As String
    withoutTechlevel = mapNameWithoutTechlevel(name_)
    formatMapName = withoutTechlevel & " (" & CStr(game.maxPlayers) & "P)"
End Function

' MARK: - WzMultiLobbyOptionsImpl

Public Sub WzMultiLobbyOptionsImpl_construct(ByVal self As WIDGET)
    Dim d As WzMultiLobbyOptionsImpl
    WzMultiLobbyOptionsWidgetBase_construct self
    self.kind = K_WzMultiLobbyOptionsImpl ' (VB) virtual calls from here on reach WzMultiLobbyOptionsImpl's overrides
    Set d = WzMultiLobbyOptionsImpl_(self)
    Set d.multiOptionButtons = New Collection
    Set d.lobbyOptionButtonUpdateFuncs = New Collection
    d.paddingXLeft = 13
    d.paddingXRight = 10
    d.paddingY = 3
    d.isChallenge = False
    d.has_showTotalVotes = False     ' = nullopt
    d.weakSerial = WeakWidget_register(self)
End Sub

Public Sub WzMultiLobbyOptionsImpl_destruct(ByVal self As WIDGET)
    Dim d As WzMultiLobbyOptionsImpl
    Set d = WzMultiLobbyOptionsImpl_(self)
    If Not d.optionsPopover Is Nothing Then
        PopoverWidget_close d.optionsPopover
    End If
    WeakWidget_unregister d.weakSerial
End Sub

' static std::shared_ptr<WzMultiLobbyOptionsImpl> make(bool isChallenge, const std::shared_ptr<WzMultiplayerOptionsTitleUI>& titleUI)
Public Function WzMultiLobbyOptionsImpl_make(ByVal isChallenge As Boolean, ByVal titleUI As WzMultiplayerOptionsTitleUI) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzMultiLobbyOptionsImpl)
    WzMultiLobbyOptionsImpl_construct widget
    WzMultiLobbyOptionsImpl_initialize widget, isChallenge, titleUI
    Set WzMultiLobbyOptionsImpl_make = widget
End Function

Private Sub displayWrappedMainRoomItem(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, w As Long, h As Long
    x0 = xOffset + psWidget.x()
    y0 = yOffset + psWidget.y()
    w = psWidget.width()
    h = psWidget.height()
    pie_BoxFill x0, y0, x0 + w, y0 + h, psPalette(WZCOL_MENU_BACKGROUND)
End Sub

Private Sub displayWrappedOptionSectionItem(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, w As Long, h As Long
    x0 = xOffset + psWidget.x()
    y0 = yOffset + psWidget.y()
    w = psWidget.width()
    h = psWidget.height()
    pie_UniTransBoxFill x0, y0, x0 + w, y0 + h, pal_RGBA(14, 12, 55, 255)
End Sub

Private Sub displayInlineEditBox(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    If W_EDITBOX_isEditing(psWidget) Then
        drawBlueBox x, y, psWidget.width(), psWidget.height()
    ElseIf (W_EDITBOX_(psWidget).state And WEDBS_HILITE) <> 0 Then
        iV_Box x, y, x + psWidget.width(), y + psWidget.height(), psPalette(WZCOL_FORM_HILITE)
    End If
End Sub

Public Function makeSectionLabelWidget(ByVal sectionName As String) As WIDGET
    Dim result As WIDGET, sectionImgColor As Long
    Set result = WzFrontendImageButton_make(True, IMAGE_CARET_DOWN_FILL)
    WzFrontendImageButton_setImageDimensions result, 12
    result.setString sectionName
    WzFrontendImageButton_setPadding result, 4, 0
    WzFrontendImageButton_setImageHorizontalOffset result, -10
    WzFrontendImageButton_setBorderDrawMode result, BorderDrawMode_Never
    result.setTransparentToMouse True
    WzFrontendImageButton_setCustomTextColors result, True, psPalette(WZCOL_FORM_TEXT), True, psPalette(WZCOL_FORM_TEXT)
    sectionImgColor = psPalette(WZCOL_TEXT_MEDIUM)
    sectionImgColor = pielightWithA(sectionImgColor, pielightA(sectionImgColor) \ 2)
    WzFrontendImageButton_setCustomImageColor result, True, sectionImgColor
    Set makeSectionLabelWidget = result
End Function

Private Sub gameNameEditHandlerFunc(ByVal widg As WIDGET)
    Dim str_ As String, priorGameName As String
    str_ = widg.getString()
    If Len(str_) = 0 Then
        widg.setString game.name_
        Exit Sub
    End If

    priorGameName = game.name_
    game.name_ = Left$(str_, 127)      ' sstrcpy(game.name, str.toUtf8().c_str()): char name[128]
    removeWildcards game.name_
    widg.setString game.name_

    If (priorGameName <> str_) And NetPlay_isHost And NetPlay_bComms Then
        NETsetLobbyOptField game.name_, NET_LOBBY_OPT_FIELD_GNAME
        sendOptions

        displayRoomSystemMessage "Game Name Updated."
    End If
End Sub

' initialize's auto wrapItemForList = [&](const std::shared_ptr<WIDGET>& widg, WIDGET_DISPLAY displayFunc = nullptr) {...}
Private Function WzMultiLobbyOptionsImpl_wrapItemForList(ByVal self As WIDGET, ByVal widg As WIDGET, Optional ByVal displayFunc As WzFunc = Nothing) As WIDGET
    Dim d As WzMultiLobbyOptionsImpl, wrappedWidget As WIDGET
    Set d = WzMultiLobbyOptionsImpl_(self)
    Set wrappedWidget = Margin_wrap(d.paddingY, d.paddingXRight, d.paddingY, d.paddingXLeft, widg)
    wrappedWidget.setGeometry 0, 0, wrappedWidget.idealWidth(), wrappedWidget.idealHeight()
    Set wrappedWidget.displayFunction = displayFunc
    Set WzMultiLobbyOptionsImpl_wrapItemForList = wrappedWidget
End Function

' initialize's auto addSectionImageButton = [&](const std::shared_ptr<LobbyMultiOptionsSectionWidget>& widg, UWORD frontendImgId, int value) -> std::shared_ptr<WzFrontendImageButton> {...}
Private Function addSectionImageButton(ByVal widg As WIDGET, ByVal frontendImgId As Long, ByVal value As Long) As WIDGET
    Dim but As WIDGET
    Set but = WzFrontendImageButton_make(True, frontendImgId)
    WzFrontendImageButton_setBorderDrawMode but, BorderDrawMode_Highlighted
    WzFrontendImageButton_setPadding but, 0, 0
    WzFrontendImageButton_setImageDimensions but, 16
    but.setGeometry 0, 0, MultiOptionsRowHeight, MultiOptionsRowHeight
    MultibuttonWidget_addButton widg, value, but
    Set addSectionImageButton = but
End Function

Private Sub WzMultiLobbyOptionsImpl_initialize(ByVal self As WIDGET, ByVal isChallenge_ As Boolean, ByVal titleUI As WzMultiplayerOptionsTitleUI)
    Dim d As WzMultiLobbyOptionsImpl, psWeakMultiLobbyOptions As Long, psWeakTitleUI As Long, locked As MultiplayOptionsLocked
    Dim optionIsHost As WzFunc, listBackgroundColor As Long, a As Long, roomNameWidget As WIDGET, psWeakGameNameEditBox As Long
    Dim roomNameEditButton As WIDGET, hideMapRow As Boolean, mapWidget As WIDGET, mapViewButton As WIDGET, mapChangeButton As WIDGET
    Dim matchOptionsSection As WIDGET, matchOptionsLabel As WIDGET, matchOptionsEditMode As Boolean
    Dim limitsSection As WIDGET, limitsLabel As WIDGET, limitsInfoBut As WIDGET, limitsOptionsBut As WIDGET, h As Long
    Set d = WzMultiLobbyOptionsImpl_(self)
    d.isChallenge = isChallenge_

    psWeakMultiLobbyOptions = d.weakSerial
    psWeakTitleUI = ObjPtr(titleUI)

    locked = getLockedOptions()

    Set d.lobbyStatusBar = WzMultiLobbyStatusBar_make(d.isChallenge)
    self.attach d.lobbyStatusBar

    Set optionIsHost = WzFunc_new(FN_MLO_OPTION_IS_HOST)

    ' (wrapItemForList, addSectionImageButton: the Private procedures above)

    Set d.optionsList = ScrollableListWidget_make()
    listBackgroundColor = psPalette(WZCOL_TRANSPARENT_BOX)
    a = pielightA(listBackgroundColor)
    If a < 200 Then a = 200
    listBackgroundColor = pielightWithA(listBackgroundColor, a)
    ScrollableListWidget_setBackgroundColor d.optionsList, listBackgroundColor
    ScrollableListWidget_setItemSpacing d.optionsList, 1
    ScrollableListWidget_setDrawRowLines d.optionsList, True

    ' Add "Room Name" row (only in multiplayer or challenge)
    Set roomNameWidget = LobbyMultiOptionsSectionWidget_new()
    MultibuttonWidget_setOuterPaddingX roomNameWidget, -(MultibuttonWidget_getGap(roomNameWidget) + 1), 0
    Set d.gameNameEditBox = W_EDITBOX_new()
    Set W_EDITBOX_(d.gameNameEditBox).pBoxDisplay = WzFunc_new(FN_displayInlineEditBox)
    W_EDITBOX_setPlaceholder d.gameNameEditBox, "Set Game Name"
    W_EDITBOX_setPlaceholderTextColor d.gameNameEditBox, True, psPalette(WZCOL_RED)
    If Not NetPlay_bComms Then
        If d.isChallenge Then d.gameNameEditBox.setString game.name_ Else d.gameNameEditBox.setString "One-Player Skirmish"
        d.gameNameEditBox.setState WEDBS_DISABLE ' disable for one-player skirmish
    End If
    W_EDITBOX_setMaxStringSize d.gameNameEditBox, 128
    W_EDITBOX_setOnReturnHandler d.gameNameEditBox, WzFunc_new(FN_gameNameEditHandlerFunc)
    W_EDITBOX_setOnEscapeHandler d.gameNameEditBox, WzFunc_new(FN_MLO_GAMENAME_ESCAPE)
    W_EDITBOX_setOnEditingStoppedHandler d.gameNameEditBox, WzFunc_new(FN_gameNameEditHandlerFunc)
    MultibuttonWidget_setLabel_widget roomNameWidget, d.gameNameEditBox
    psWeakGameNameEditBox = ObjPtr(d.gameNameEditBox)
    If NetPlay_bComms Then
        Set roomNameEditButton = addSectionImageButton(roomNameWidget, IMAGE_PENCIL_SQUARE, 1)
        W_BUTTON_addOnClickHandler roomNameEditButton, WzFunc_new(FN_MLO_ROOMNAME_EDIT_CLICK, psWeakGameNameEditBox)
        Select Case ingame.side
            Case InGameSide_HOST_OR_SINGLEPLAYER
                roomNameEditButton.setTip "Edit Game Name"
            Case InGameSide_MULTIPLAYER_CLIENT
                roomNameEditButton.setTip "Game Name"
                roomNameEditButton.setState WBUT_DISABLE
        End Select
    End If
    If NetPlay_bComms And ingame.side = InGameSide_HOST_OR_SINGLEPLAYER Then
        If NETGameIsLocked() Then
            Set d.roomPasswordOptionsButton = addSectionImageButton(roomNameWidget, IMAGE_LOCK, 2)
        Else
            Set d.roomPasswordOptionsButton = addSectionImageButton(roomNameWidget, IMAGE_UNLOCK, 2)
        End If
        W_BUTTON_addOnClickHandler d.roomPasswordOptionsButton, WzFunc_new(FN_MLO_PASSWORD_BUTTON_CLICK, psWeakMultiLobbyOptions)
        roomNameWidget.setGeometry 0, 0, roomNameWidget.idealWidth(), roomNameWidget.idealHeight()
    End If
    ScrollableListWidget_addItem d.optionsList, WzMultiLobbyOptionsImpl_wrapItemForList(self, roomNameWidget, WzFunc_new(FN_displayWrappedMainRoomItem))

    hideMapRow = isBlindSimpleLobby(game.blindMode) And Not NetPlay_isHost
    If Not hideMapRow Then
        ' Add "Map" row
        Set mapWidget = LobbyMultiOptionsSectionWidget_new()
        Set d.mapNameWidget = W_LABEL_new() ' TODO: Replace with custom widget!!
        d.mapNameWidget.setString formatMapName(game.map)
        MultibuttonWidget_setLabel_widget mapWidget, d.mapNameWidget
        Set mapViewButton = addSectionImageButton(mapWidget, IMAGE_EYE, 1)
        mapViewButton.setTip "Click to see Map"
        W_BUTTON_addOnClickHandler mapViewButton, WzFunc_new(FN_MLO_MAPVIEW_CLICK)
        Set mapChangeButton = addSectionImageButton(mapWidget, IMAGE_GLOBE, 2)
        mapChangeButton.setTip "Select Map" & vbLf & "Can be blocked by players' votes"
        If Not titleUI Is Nothing Then
            W_BUTTON_addOnClickHandler mapChangeButton, WzFunc_new(FN_MLO_MAPCHANGE_CLICK, psWeakTitleUI)
        End If
        If d.isChallenge Or ingame.side = InGameSide_MULTIPLAYER_CLIENT Or titleUI Is Nothing Then
            mapChangeButton.setState WBUT_DISABLE
        End If
        mapWidget.setGeometry 0, 0, mapWidget.idealWidth(), mapWidget.idealHeight()
        ScrollableListWidget_addItem d.optionsList, WzMultiLobbyOptionsImpl_wrapItemForList(self, mapWidget, WzFunc_new(FN_displayWrappedMainRoomItem))
    End If

    ' SECTION: "Match Options"
    Set matchOptionsSection = LobbyMultiOptionsSectionWidget_new()
    Set matchOptionsLabel = makeSectionLabelWidget("Match Options")
    MultibuttonWidget_setLabel_widget matchOptionsSection, matchOptionsLabel
    ' Add right-side button(s)
    Set d.matchOptionsConfigureButton = addSectionImageButton(matchOptionsSection, IMAGE_GEAR, 1)
    W_BUTTON_addOnClickHandler d.matchOptionsConfigureButton, WzFunc_new(FN_MLO_MATCHOPTIONS_CLICK, psWeakMultiLobbyOptions)
    If d.isChallenge Then
        d.matchOptionsConfigureButton.setState WBUT_DISABLE
    End If
    matchOptionsSection.setGeometry 0, 0, matchOptionsSection.idealWidth(), matchOptionsSection.idealHeight()
    ScrollableListWidget_addItem d.optionsList, WzMultiLobbyOptionsImpl_wrapItemForList(self, matchOptionsSection, WzFunc_new(FN_displayWrappedOptionSectionItem))

    matchOptionsEditMode = (ingame.side = InGameSide_HOST_OR_SINGLEPLAYER)

    ' Scavengers
    Set d.scavengerChoice = LobbyMultiOptionsChoiceWidget_new(matchOptionsEditMode)
    MultibuttonWidget_setLabel d.scavengerChoice, "Scavengers"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.scavengerChoice, MULTIOPT_T_ScavType, ULTIMATE_SCAVENGERS, AtlasImage_(FrontImages, IMAGE_SCAVENGERS_ULTIMATE_ON), AtlasImage_(FrontImages, IMAGE_SCAVENGERS_ULTIMATE_ON_HI), "Ultimate Scavengers"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.scavengerChoice, MULTIOPT_T_ScavType, SCAVENGERS, AtlasImage_(FrontImages, IMAGE_SCAVENGERS_ON), AtlasImage_(FrontImages, IMAGE_SCAVENGERS_ON_HI), "Scavengers"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.scavengerChoice, MULTIOPT_T_ScavType, NO_SCAVENGERS, AtlasImage_(FrontImages, IMAGE_SCAVENGERS_OFF), AtlasImage_(FrontImages, IMAGE_SCAVENGERS_OFF_HI), "No Scavengers"
    MultibuttonWidget_enable d.scavengerChoice, (ingame.side = InGameSide_MULTIPLAYER_CLIENT) Or Not locked.scavengers
    d.scavengerChoice.setGeometry 0, 0, d.scavengerChoice.idealWidth(), d.scavengerChoice.idealHeight()
    MultibuttonWidget_setCanChooseHandler d.scavengerChoice, WzFunc_new(FN_MLO_SCAV_CANCHOOSE)
    MultibuttonWidget_addOnChooseHandler d.scavengerChoice, WzFunc_new(FN_MLO_SCAV_ONCHOOSE)
    ScrollableListWidget_addItem d.optionsList, WzMultiLobbyOptionsImpl_wrapItemForList(self, d.scavengerChoice)

    ' Alliances
    Set d.allianceChoice = LobbyMultiOptionsChoiceWidget_new(matchOptionsEditMode)
    MultibuttonWidget_setLabel d.allianceChoice, "Alliances"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.allianceChoice, MULTIOPT_T_AllianceType, NO_ALLIANCES, AtlasImage_(FrontImages, IMAGE_NOALLI), AtlasImage_(FrontImages, IMAGE_NOALLI_HI), "No Alliances"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.allianceChoice, MULTIOPT_T_AllianceType, ALLIANCES_, AtlasImage_(FrontImages, IMAGE_ALLI), AtlasImage_(FrontImages, IMAGE_ALLI_HI), "Allow Alliances"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.allianceChoice, MULTIOPT_T_AllianceType, ALLIANCES_UNSHARED, AtlasImage_(FrontImages, IMAGE_ALLI_UNSHARED), AtlasImage_(FrontImages, IMAGE_ALLI_UNSHARED_HI), "Locked Teams, No Shared Research"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.allianceChoice, MULTIOPT_T_AllianceType, ALLIANCES_TEAMS, AtlasImage_(FrontImages, IMAGE_ALLI_TEAMS), AtlasImage_(FrontImages, IMAGE_ALLI_TEAMS_HI), "Locked Teams"
    MultibuttonWidget_enable d.allianceChoice, (ingame.side = InGameSide_MULTIPLAYER_CLIENT) Or Not locked.alliances
    d.allianceChoice.setGeometry 0, 0, d.allianceChoice.idealWidth(), d.allianceChoice.idealHeight()
    MultibuttonWidget_setCanChooseHandler d.allianceChoice, WzFunc_new(FN_MLO_ALLIANCE_CANCHOOSE)
    MultibuttonWidget_addOnChooseHandler d.allianceChoice, WzFunc_new(FN_MLO_ALLIANCE_ONCHOOSE)
    ScrollableListWidget_addItem d.optionsList, WzMultiLobbyOptionsImpl_wrapItemForList(self, d.allianceChoice)

    ' Power
    Set d.powerChoice = LobbyMultiOptionsChoiceWidget_new(matchOptionsEditMode)
    MultibuttonWidget_setLabel d.powerChoice, "Power"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.powerChoice, MULTIOPT_T_PowerSetting, LEV_LOW, AtlasImage_(FrontImages, IMAGE_POWLO), AtlasImage_(FrontImages, IMAGE_POWLO_HI), "Low Power Levels"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.powerChoice, MULTIOPT_T_PowerSetting, LEV_MED, AtlasImage_(FrontImages, IMAGE_POWMED), AtlasImage_(FrontImages, IMAGE_POWMED_HI), "Medium Power Levels"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.powerChoice, MULTIOPT_T_PowerSetting, LEV_HI, AtlasImage_(FrontImages, IMAGE_POWHI), AtlasImage_(FrontImages, IMAGE_POWHI_HI), "High Power Levels"
    MultibuttonWidget_enable d.powerChoice, (ingame.side = InGameSide_MULTIPLAYER_CLIENT) Or Not locked.power
    d.powerChoice.setGeometry 0, 0, d.powerChoice.idealWidth(), d.powerChoice.idealHeight()
    MultibuttonWidget_setCanChooseHandler d.powerChoice, WzFunc_new(FN_MLO_POWER_CANCHOOSE)
    MultibuttonWidget_addOnChooseHandler d.powerChoice, WzFunc_new(FN_MLO_POWER_ONCHOOSE)
    ScrollableListWidget_addItem d.optionsList, WzMultiLobbyOptionsImpl_wrapItemForList(self, d.powerChoice)

    ' Bases
    Set d.baseTypeChoice = LobbyMultiOptionsChoiceWidget_new(matchOptionsEditMode)
    MultibuttonWidget_setLabel d.baseTypeChoice, "Base"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.baseTypeChoice, MULTIOPT_T_CampType, CAMP_CLEAN, AtlasImage_(FrontImages, IMAGE_NOBASE), AtlasImage_(FrontImages, IMAGE_NOBASE_HI), "Start with No Bases"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.baseTypeChoice, MULTIOPT_T_CampType, CAMP_BASE, AtlasImage_(FrontImages, IMAGE_SBASE), AtlasImage_(FrontImages, IMAGE_SBASE_HI), "Start with Bases"
    WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.baseTypeChoice, MULTIOPT_T_CampType, CAMP_WALLS, AtlasImage_(FrontImages, IMAGE_LBASE), AtlasImage_(FrontImages, IMAGE_LBASE_HI), "Start with Advanced Bases"
    MultibuttonWidget_enable d.baseTypeChoice, (ingame.side = InGameSide_MULTIPLAYER_CLIENT) Or Not locked.bases
    d.baseTypeChoice.setGeometry 0, 0, d.baseTypeChoice.idealWidth(), d.baseTypeChoice.idealHeight()
    MultibuttonWidget_setCanChooseHandler d.baseTypeChoice, WzFunc_new(FN_MLO_BASE_CANCHOOSE)
    MultibuttonWidget_addOnChooseHandler d.baseTypeChoice, WzFunc_new(FN_MLO_BASE_ONCHOOSE)
    ScrollableListWidget_addItem d.optionsList, WzMultiLobbyOptionsImpl_wrapItemForList(self, d.baseTypeChoice)

    If Not d.isChallenge Then
        ' Tech Level
        Set d.technologyChoice = LobbyMultiOptionsChoiceWidget_new(matchOptionsEditMode)
        MultibuttonWidget_setLabel d.technologyChoice, "Tech"
        WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.technologyChoice, MULTIOPT_T_TechLevel, TECH_1, AtlasImage_(FrontImages, IMAGE_TECHLO), AtlasImage_(FrontImages, IMAGE_TECHLO_HI), "Technology Level 1"
        WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.technologyChoice, MULTIOPT_T_TechLevel, TECH_2, AtlasImage_(FrontImages, IMAGE_TECHMED), AtlasImage_(FrontImages, IMAGE_TECHMED_HI), "Technology Level 2"
        WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.technologyChoice, MULTIOPT_T_TechLevel, TECH_3, AtlasImage_(FrontImages, IMAGE_TECHHI), AtlasImage_(FrontImages, IMAGE_TECHHI_HI), "Technology Level 3"
        WzMultiLobbyOptionsImpl_addMultiOptButtonT self, d.technologyChoice, MULTIOPT_T_TechLevel, TECH_4, AtlasImage_(FrontImages, IMAGE_COMPUTER_Y), AtlasImage_(FrontImages, IMAGE_COMPUTER_Y_HI), "Technology Level 4"
        d.technologyChoice.setGeometry 0, 0, d.technologyChoice.idealWidth(), d.technologyChoice.idealHeight()
        MultibuttonWidget_setCanChooseHandler d.technologyChoice, optionIsHost
        MultibuttonWidget_addOnChooseHandler d.technologyChoice, WzFunc_new(FN_MLO_TECH_ONCHOOSE)
        ScrollableListWidget_addItem d.optionsList, WzMultiLobbyOptionsImpl_wrapItemForList(self, d.technologyChoice)
    End If

    ' Refresh values for match options
    WzMultiLobbyOptionsImpl_setGameOptionStates self

    ' SECTION: "Limits"
    Set limitsSection = LobbyMultiOptionsSectionWidget_new()
    Set limitsLabel = makeSectionLabelWidget("Limits")
    MultibuttonWidget_setLabel_widget limitsSection, limitsLabel
    Set limitsInfoBut = addSectionImageButton(limitsSection, IMAGE_INFO_CIRCLE, 1)
    W_BUTTON_addOnClickHandler limitsInfoBut, WzFunc_new(FN_MLO_LIMITSINFO_CLICK)
    If ingame.side = InGameSide_HOST_OR_SINGLEPLAYER Then
        Set limitsOptionsBut = addSectionImageButton(limitsSection, IMAGE_GEAR, 2)
        W_BUTTON_addOnClickHandler limitsOptionsBut, WzFunc_new(FN_MLO_LIMITSOPTIONS_CLICK)
    End If
    limitsSection.setGeometry 0, 0, limitsSection.idealWidth(), limitsSection.idealHeight()
    ScrollableListWidget_addItem d.optionsList, WzMultiLobbyOptionsImpl_wrapItemForList(self, limitsSection, WzFunc_new(FN_displayWrappedOptionSectionItem))

    Set d.limitIconsWidget = LobbyMultiOptionsSectionWidget_new()
    WzMultiLobbyOptionsImpl_addLimitIcons self, d.limitIconsWidget
    h = d.limitIconsWidget.idealHeight()
    If h < MultiOptionsRowHeight Then h = MultiOptionsRowHeight
    d.limitIconsWidget.setGeometry 0, 0, d.limitIconsWidget.idealWidth(), h
    Set d.limitIconsWidgetWrapper = WzMultiLobbyOptionsImpl_wrapItemForList(self, d.limitIconsWidget)
    d.limitIconsWidgetWrapper.show MultibuttonWidget_numButtons(d.limitIconsWidget) > 0
    ScrollableListWidget_addItem d.optionsList, d.limitIconsWidgetWrapper

    self.attach d.optionsList
End Sub

' template<typename T> void addMultiOptButtonT(const std::shared_ptr<MultibuttonWidget>& mbw, T pref, AtlasImage image, AtlasImage imageDown, char const *tip)
' (prefType: T, a MULTIOPT_T_ value)
Private Sub WzMultiLobbyOptionsImpl_addMultiOptButtonT(ByVal self As WIDGET, ByVal mbw As WIDGET, ByVal prefType As Long, ByVal pref As Long, image As AtlasImage, imageDown As AtlasImage, ByVal tip As String)
    Dim d As WzMultiLobbyOptionsImpl, button As WIDGET, showLocalPrefValue As Boolean, weakSelf As Long, updateDataFunc As WzFunc
    Dim highlightImage As AtlasImage, nullImage As AtlasImage
    Set d = WzMultiLobbyOptionsImpl_(self)
    Set button = WzLobbyMultiOptionButton_new()
    highlightImage = mpwidgetGetFrontHighlightImage(image)
    nullImage = AtlasImage_()
    W_BUTTON_setImages_atlas button, image, imageDown, highlightImage, nullImage
    button.setTip tip

    showLocalPrefValue = (ingame.side = InGameSide_MULTIPLAYER_CLIENT) And Not (NetPlay_isHost)

    If showLocalPrefValue Then
        WzLobbyMultiOptionButton_setMode button, MultiOptionButtonMode_CLIENT
        W_BUTTON_(button).minClickInterval = GAME_TICKS_PER_SEC
    End If
    WzLobbyMultiOptionButton_setShowTotalVotes button, (ingame.side = InGameSide_HOST_OR_SINGLEPLAYER) And NetPlay_isHost

    weakSelf = ObjPtr(button)
    Set updateDataFunc = WzFunc_new(FN_MLO_OPTBUTTON_UPDATE, prefType, pref, weakSelf)
    WzFunc_call updateDataFunc

    If showLocalPrefValue Then
        W_BUTTON_addOnClickHandler button, WzFunc_new(FN_MLO_OPTBUTTON_CLICK, prefType, pref)
        WzLobbyMultiOptionButton_setShowLocalVoteIndicatorOnHighlight button, True
    End If

    MultibuttonWidget_addButton mbw, pref, button
    d.multiOptionButtons.Add button

    d.lobbyOptionButtonUpdateFuncs.Add updateDataFunc
End Sub

' (VB) multivote's overloads, by addMultiOptButtonT's T
Private Function getMultiOptionPrefValueT(ByVal prefType As Long, ByVal val_ As Long) As Boolean
    Select Case prefType
    Case MULTIOPT_T_ScavType: getMultiOptionPrefValueT = getMultiOptionPrefValue_ScavType(val_)
    Case MULTIOPT_T_AllianceType: getMultiOptionPrefValueT = getMultiOptionPrefValue_AllianceType(val_)
    Case MULTIOPT_T_PowerSetting: getMultiOptionPrefValueT = getMultiOptionPrefValue_PowerSetting(val_)
    Case MULTIOPT_T_CampType: getMultiOptionPrefValueT = getMultiOptionPrefValue_CampType(val_)
    Case MULTIOPT_T_TechLevel: getMultiOptionPrefValueT = getMultiOptionPrefValue_TechLevel(val_)
    End Select
End Function

Private Sub setMultiOptionPrefValueT(ByVal prefType As Long, ByVal val_ As Long, ByVal pref As Boolean)
    Select Case prefType
    Case MULTIOPT_T_ScavType: setMultiOptionPrefValue_ScavType val_, pref
    Case MULTIOPT_T_AllianceType: setMultiOptionPrefValue_AllianceType val_, pref
    Case MULTIOPT_T_PowerSetting: setMultiOptionPrefValue_PowerSetting val_, pref
    Case MULTIOPT_T_CampType: setMultiOptionPrefValue_CampType val_, pref
    Case MULTIOPT_T_TechLevel: setMultiOptionPrefValue_TechLevel val_, pref
    End Select
End Sub

Private Function getMultiOptionPrefValueTotalT(ByVal prefType As Long, ByVal val_ As Long) As Long
    Select Case prefType
    Case MULTIOPT_T_ScavType: getMultiOptionPrefValueTotalT = getMultiOptionPrefValueTotal_ScavType(val_)
    Case MULTIOPT_T_AllianceType: getMultiOptionPrefValueTotalT = getMultiOptionPrefValueTotal_AllianceType(val_)
    Case MULTIOPT_T_PowerSetting: getMultiOptionPrefValueTotalT = getMultiOptionPrefValueTotal_PowerSetting(val_)
    Case MULTIOPT_T_CampType: getMultiOptionPrefValueTotalT = getMultiOptionPrefValueTotal_CampType(val_)
    Case MULTIOPT_T_TechLevel: getMultiOptionPrefValueTotalT = getMultiOptionPrefValueTotal_TechLevel(val_)
    End Select
End Function

' initialize's ... setGameOptionStates' auto setButtonState = [](const std::shared_ptr<LobbyMultiOptionsChoiceWidget>& mbw, int value, int state) {...}
Private Sub setButtonState(ByVal mbw As WIDGET, ByVal value As Long, ByVal state As Long)
    Dim pButton As WIDGET
    Set pButton = MultibuttonWidget_getButtonByValue(mbw, value)
    If Not pButton Is Nothing Then
        pButton.setState state
    End If
End Sub

Private Sub WzMultiLobbyOptionsImpl_setGameOptionStates(ByVal self As WIDGET)
    Dim d As WzMultiLobbyOptionsImpl, locked As MultiplayOptionsLocked
    Set d = WzMultiLobbyOptionsImpl_(self)
    locked = getLockedOptions()

    If NetPlay_bComms Then
        d.gameNameEditBox.setString game.name_
        If ingame.side = InGameSide_MULTIPLAYER_CLIENT Then
            d.gameNameEditBox.setState WEDBS_DISABLE
        End If
    Else
        If d.isChallenge Then d.gameNameEditBox.setString game.name_ Else d.gameNameEditBox.setString "One-Player Skirmish"
        d.gameNameEditBox.setState WEDBS_DISABLE ' disable for one-player skirmish
    End If

    If Not d.roomPasswordOptionsButton Is Nothing Then
        If NETGameIsLocked() Then
            WzFrontendImageButton_setImage d.roomPasswordOptionsButton, True, IMAGE_LOCK
        Else
            WzFrontendImageButton_setImage d.roomPasswordOptionsButton, True, IMAGE_UNLOCK
        End If
    End If

    If Not d.mapNameWidget Is Nothing Then
        ' Update map name & details
        d.mapNameWidget.setString formatMapName(game.map)
    End If

    ' (setButtonState: the Private procedure above)

    If Not game.mapHasScavengers Then
        setButtonState d.scavengerChoice, ULTIMATE_SCAVENGERS, WBUT_DISABLE
        setButtonState d.scavengerChoice, SCAVENGERS, WBUT_DISABLE
    ElseIf Not MultibuttonWidget_isDisabled(d.scavengerChoice) Then
        setButtonState d.scavengerChoice, ULTIMATE_SCAVENGERS, 0
        setButtonState d.scavengerChoice, SCAVENGERS, 0
    End If

    ' Update MultiButtonWidget enabled status
    MultibuttonWidget_enable d.scavengerChoice, (ingame.side = InGameSide_MULTIPLAYER_CLIENT) Or Not locked.scavengers
    MultibuttonWidget_enable d.allianceChoice, (ingame.side = InGameSide_MULTIPLAYER_CLIENT) Or Not locked.alliances
    MultibuttonWidget_enable d.powerChoice, (ingame.side = InGameSide_MULTIPLAYER_CLIENT) Or Not locked.power
    MultibuttonWidget_enable d.baseTypeChoice, (ingame.side = InGameSide_MULTIPLAYER_CLIENT) Or Not locked.bases

    MultibuttonWidget_choose d.scavengerChoice, game.scavengers
    MultibuttonWidget_choose d.allianceChoice, game.alliance
    MultibuttonWidget_choose d.powerChoice, game.power
    MultibuttonWidget_choose d.baseTypeChoice, game.base
    If Not d.technologyChoice Is Nothing Then
        MultibuttonWidget_choose d.technologyChoice, game.techLevel
    End If
End Sub

Public Function WzMultiLobbyOptionsImpl_getShowTotalOptionVotes(ByVal self As WIDGET) As Boolean
    Dim d As WzMultiLobbyOptionsImpl
    Set d = WzMultiLobbyOptionsImpl_(self)
    If d.has_showTotalVotes Then
        WzMultiLobbyOptionsImpl_getShowTotalOptionVotes = d.showTotalVotes
    Else
        WzMultiLobbyOptionsImpl_getShowTotalOptionVotes = (ingame.side = InGameSide_HOST_OR_SINGLEPLAYER) And NetPlay_isHost
    End If
End Function

Public Sub WzMultiLobbyOptionsImpl_refreshData(ByVal self As WIDGET)
    Dim d As WzMultiLobbyOptionsImpl, i As Long, buttonUpdateFunc As WzFunc
    Set d = WzMultiLobbyOptionsImpl_(self)
    WzMultiLobbyStatusBar_refreshData d.lobbyStatusBar

    WzMultiLobbyOptionsImpl_setGameOptionStates self

    ' Update individual option buttons
    For i = 1 To d.lobbyOptionButtonUpdateFuncs.Count
        Set buttonUpdateFunc = d.lobbyOptionButtonUpdateFuncs(i)
        WzFunc_call buttonUpdateFunc
    Next
    WzMultiLobbyOptionsImpl_setShowTotalOptionVotesImpl self, WzMultiLobbyOptionsImpl_getShowTotalOptionVotes(self)

    ' update limit icons
    MultibuttonWidget_clearButtons d.limitIconsWidget
    WzMultiLobbyOptionsImpl_addLimitIcons self, d.limitIconsWidget
    d.limitIconsWidgetWrapper.show MultibuttonWidget_numButtons(d.limitIconsWidget) > 0

    ScrollableListWidget_setLayoutDirty d.optionsList
End Sub

Public Sub WzMultiLobbyOptionsImpl_forceSetShowTotalOptionVotes(ByVal self As WIDGET, ByVal show As Boolean)
    Dim d As WzMultiLobbyOptionsImpl
    Set d = WzMultiLobbyOptionsImpl_(self)
    d.has_showTotalVotes = True
    d.showTotalVotes = show
    WzMultiLobbyOptionsImpl_setShowTotalOptionVotesImpl self, d.showTotalVotes
End Sub

Private Sub WzMultiLobbyOptionsImpl_setShowTotalOptionVotesImpl(ByVal self As WIDGET, ByVal show As Boolean)
    Dim d As WzMultiLobbyOptionsImpl, i As Long, button As WIDGET
    Set d = WzMultiLobbyOptionsImpl_(self)
    For i = 1 To d.multiOptionButtons.Count
        Set button = d.multiOptionButtons(i)
        WzLobbyMultiOptionButton_setShowTotalVotes button, show
    Next
End Sub

Private Sub WzMultiLobbyOptionsImpl_addLimitIcons(ByVal self As WIDGET, ByVal mbw As WIDGET)
    Dim i As Long, button As WIDGET, icon As AtlasImage
    For i = 0 To limitIcons_size() - 1
        If (ingame.flags And SHL(1, i)) <> 0 Then
            Set button = WzLobbyMultiOptionButton_new()
            icon = AtlasImage_(FrontImages, limitIcons(i).icon)
            W_BUTTON_setImages_atlas button, icon, icon, icon, AtlasImage_()
            button.setTip limitIcons(i).desc

            MultibuttonWidget_addButton mbw, i, button
        End If
    Next
End Sub

Public Function WzMultiLobbyOptionsImpl_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WzMultiLobbyOptionsImpl, a As Long, b As Long
    Set d = WzMultiLobbyOptionsImpl_(self)
    a = d.lobbyStatusBar.idealWidth()
    b = d.optionsList.idealWidth()
    If b > a Then a = b
    WzMultiLobbyOptionsImpl_idealWidth = a
End Function

Public Function WzMultiLobbyOptionsImpl_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzMultiLobbyOptionsImpl, result As Long
    Set d = WzMultiLobbyOptionsImpl_(self)
    result = d.lobbyStatusBar.idealHeight()
    result = result + d.optionsList.idealHeight()
    result = result + 1
    WzMultiLobbyOptionsImpl_idealHeight = result
End Function

Public Sub WzMultiLobbyOptionsImpl_geometryChanged(ByVal self As WIDGET)
    Dim d As WzMultiLobbyOptionsImpl, w As Long, h As Long, wrappedOptionRowHeight As Long, idealListHeight As Long
    Dim availableListHeight As Long, listY0 As Long
    Set d = WzMultiLobbyOptionsImpl_(self)
    w = self.width()
    h = self.height()

    If w = 0 Or h = 0 Then
        Exit Sub
    End If

    d.lobbyStatusBar.setGeometry 0, 0, w, d.lobbyStatusBar.idealHeight()

    wrappedOptionRowHeight = d.paddingY + MultiOptionsRowHeight + d.paddingY
    idealListHeight = d.optionsList.idealHeight()
    availableListHeight = h - d.lobbyStatusBar.height() - 1
    If idealListHeight < availableListHeight Then availableListHeight = idealListHeight
    If availableListHeight < idealListHeight Then
        availableListHeight = (availableListHeight \ (wrappedOptionRowHeight + 1)) * (wrappedOptionRowHeight + 1)
    End If
    listY0 = d.lobbyStatusBar.height()
    d.optionsList.setGeometry 0, listY0, w, availableListHeight
End Sub

' displayOptionsOverlay(const std::shared_ptr<WIDGET>& psParent, const std::shared_ptr<WIDGET>& optionsPopOver, const OnOverlayCloseHandlerFunc& closeHandler = nullptr)
Private Sub WzMultiLobbyOptionsImpl_displayOptionsOverlay(ByVal self As WIDGET, ByVal psParent As WIDGET, ByVal optionsPopOver As WIDGET, Optional ByVal closeHandler As WzFunc = Nothing)
    Dim d As WzMultiLobbyOptionsImpl
    Set d = WzMultiLobbyOptionsImpl_(self)
    If Not d.optionsPopover Is Nothing Then
        PopoverWidget_close d.optionsPopover
    End If
    Set d.optionsPopover = PopoverWidget_makePopover(psParent, optionsPopOver, PopoverWidget_Style_Interactive, PopoverWidget_Alignment_LeftOfParent, 0, 1, closeHandler)
End Sub

Private Sub WzMultiLobbyOptionsImpl_displayPasswordOverlayForm(ByVal self As WIDGET, ByVal psParent As WIDGET)
    Dim d As WzMultiLobbyOptionsImpl, weakMultiLobbyOptions As Long, passwordPopover As WIDGET, weakPasswordPopover As Long
    Set d = WzMultiLobbyOptionsImpl_(self)
    weakMultiLobbyOptions = d.weakSerial
    d.roomPasswordOptionsButton.setState WBUT_CLICKLOCK

    Set passwordPopover = WzLobbyPasswordOptionPopover_make(weakMultiLobbyOptions, d.cachedPasswordFieldValue)
    passwordPopover.setGeometry 0, 0, passwordPopover.idealWidth(), passwordPopover.idealHeight()

    weakPasswordPopover = WzLobbyPasswordOptionPopover_(passwordPopover).weakSerial
    WzMultiLobbyOptionsImpl_displayOptionsOverlay self, psParent, passwordPopover, WzFunc_new(FN_MLO_PWOVERLAY_CLOSE, weakPasswordPopover, weakMultiLobbyOptions)

    widgScheduleTask WzFunc_new(FN_MLO_PWOVERLAY_FOCUS_TASK, weakPasswordPopover)
End Sub

' static std::shared_ptr<WzClickableOptionsChoiceWidget> addClickableMenuItem(const std::shared_ptr<PopoverMenuWidget>& menu, const WzString& text, const std::function<void (WIDGET_KEY)>& onClick, bool disabled = false)
Private Function addClickableMenuItem(ByVal menu As WIDGET, ByVal text As String, ByVal onClick As WzFunc, Optional ByVal disabled As Boolean = False) As WIDGET
    Dim result As WIDGET
    Set result = WzClickableOptionsChoiceWidget_make(font_regular)
    result.setString text
    WzOptionsChoiceWidget_setTextAlignment result, WLAB_ALIGNLEFT
    If Not onClick Is Nothing Then
        WzClickableOptionsChoiceWidget_addOnClickHandler result, WzFunc_new(FN_MLO_MENUITEM_CLICK, onClick)
    End If
    WzOptionsChoiceWidget_setDisabled result, disabled
    result.setGeometry 0, 0, result.idealWidth(), result.idealHeight()
    PopoverMenuWidget_addMenuItem menu, result, True
    Set addClickableMenuItem = result
End Function

Private Function WzMultiLobbyOptionsImpl_buildMatchOptionsMenu(ByVal self As WIDGET) As WIDGET
    Dim d As WzMultiLobbyOptionsImpl, weakMultiLobbyOptions As Long, popoverMenu As WIDGET, idealMenuHeight As Long, menuHeight As Long
    Set d = WzMultiLobbyOptionsImpl_(self)
    weakMultiLobbyOptions = d.weakSerial

    Set popoverMenu = PopoverMenuWidget_make()

    If ingame.side = InGameSide_HOST_OR_SINGLEPLAYER And Not d.isChallenge Then
        ' "Randomize Options"
        addClickableMenuItem popoverMenu, "Randomize Options", WzFunc_new(FN_MLO_MENU_RANDOMIZE)
    End If

    If (ingame.side = InGameSide_MULTIPLAYER_CLIENT) And Not (NetPlay_isHost) Then
        ' [Match Option Preferences:]
        addClickableMenuItem popoverMenu, "Match Option Preferences:", Nothing, True
        ' Enable All Options
        addClickableMenuItem popoverMenu, "Star All", WzFunc_new(FN_MLO_MENU_STARALL)
        ' Clear All Options
        addClickableMenuItem popoverMenu, "Clear All", WzFunc_new(FN_MLO_MENU_CLEARALL)
    End If

    If (ingame.side = InGameSide_HOST_OR_SINGLEPLAYER) And Not d.isChallenge Then
        addClickableMenuItem popoverMenu, "Match Option Preferences:", Nothing, True
        If Not WzMultiLobbyOptionsImpl_getShowTotalOptionVotes(self) Then
            addClickableMenuItem popoverMenu, "Show Player Preferences", WzFunc_new(FN_MLO_MENU_SHOWPREFS, weakMultiLobbyOptions, True)
        Else
            addClickableMenuItem popoverMenu, "Hide Player Preferences", WzFunc_new(FN_MLO_MENU_SHOWPREFS, weakMultiLobbyOptions, False)
        End If
    End If

    idealMenuHeight = popoverMenu.idealHeight()
    menuHeight = idealMenuHeight
    If menuHeight > screenHeight Then
        menuHeight = screenHeight
    End If
    popoverMenu.setGeometry popoverMenu.x(), popoverMenu.y(), popoverMenu.idealWidth(), menuHeight

    Set WzMultiLobbyOptionsImpl_buildMatchOptionsMenu = popoverMenu
End Function

Private Sub WzMultiLobbyOptionsImpl_displayMatchOptionsMenu(ByVal self As WIDGET, ByVal psParent As WIDGET)
    Dim d As WzMultiLobbyOptionsImpl, popoverMenu As WIDGET, weakMultiLobbyOptions As Long
    Set d = WzMultiLobbyOptionsImpl_(self)
    Set popoverMenu = WzMultiLobbyOptionsImpl_buildMatchOptionsMenu(self)

    weakMultiLobbyOptions = d.weakSerial
    d.matchOptionsConfigureButton.setState WBUT_CLICKLOCK

    If Not d.optionsPopover Is Nothing Then
        PopoverWidget_close d.optionsPopover
    End If
    Set d.optionsPopover = PopoverMenuWidget_openMenu(popoverMenu, psParent, PopoverWidget_Alignment_LeftOfParent, 0, 1, WzFunc_new(FN_MLO_MATCHMENU_CLOSE, weakMultiLobbyOptions))
End Sub

' std::shared_ptr<WzMultiLobbyOptionsWidgetBase> makeWzMultiLobbyOptionsForm(bool isChallenge, const std::shared_ptr<WzMultiplayerOptionsTitleUI>& titleUI)
Public Function makeWzMultiLobbyOptionsForm(ByVal isChallenge As Boolean, ByVal titleUI As WzMultiplayerOptionsTitleUI) As WIDGET
    Set makeWzMultiLobbyOptionsForm = WzMultiLobbyOptionsImpl_make(isChallenge, titleUI)
End Function

' MARK: - the lambdas and display functions

Public Function multilobbyoptions_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongPopover As WIDGET, widg As WIDGET, newPassword As String, willSet As Boolean, pd As WzLobbyPasswordOptionPopover
    Dim strongEditBox As WIDGET, context As W_CONTEXT, strongMultiLobbyOptions As WIDGET, multiOptionsTitleUI As WzMultiplayerOptionsTitleUI
    Dim newValue As Long, strongSelf As WIDGET, showLocalPrefValue As Boolean, prefValue As Boolean, showTotalPrefValues As Boolean
    Dim prefValueTotal As Long, b As WIDGET, psButton As WIDGET, newPref As Boolean, strongPasswordPopover As WIDGET
    Dim parentTitleUI As WzMultiplayerOptionsTitleUI, line_ As String
    Select Case f.id
    Case FN_displayPasswordEditBox
        displayPasswordEditBox args(0), args(1), args(2)

    Case FN_MLO_PW_RETURN                       ' [weakPopover](W_EDITBOX& widg)
        Set widg = args(0)
        Set strongPopover = WeakWidget_lock(f.cap(0))
        If Not (Not strongPopover Is Nothing) Then ASSERT_FAILURE "strongPopover != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:485", "initialize", "No parent?": Exit Function
        If Not ((NetPlay_bComms And ingame.side = InGameSide_HOST_OR_SINGLEPLAYER)) Then ASSERT_FAILURE "(NetPlay.bComms && ingame.side == InGameSide::HOST_OR_SINGLEPLAYER)", "src/titleui/widgets/multilobbyoptions.cpp:486", "initialize", "Not the host?": Exit Function
        Set pd = WzLobbyPasswordOptionPopover_(strongPopover)
        newPassword = widg.getString()
        willSet = (WzAdvCheckbox_isChecked(pd.requiredPasswordCheckbox)) Or ((pd.requiredPasswordCheckbox.getState() And WBUT_DISABLE) <> 0)
        willSet = willSet And (Len(newPassword) <> 0)
        If willSet Then
            WzLobbyPasswordOptionPopover_setGamePassword strongPopover, newPassword
        Else
            WzLobbyPasswordOptionPopover_resetGamePassword strongPopover
        End If

    Case FN_MLO_PW_ESCAPE                       ' [weakPopover](W_EDITBOX& widg)
        Set widg = args(0)
        Set strongPopover = WeakWidget_lock(f.cap(0))
        If Not (Not strongPopover Is Nothing) Then ASSERT_FAILURE "strongPopover != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:501", "initialize", "No parent?": Exit Function
        If Not ((NetPlay_bComms And ingame.side = InGameSide_HOST_OR_SINGLEPLAYER)) Then ASSERT_FAILURE "(NetPlay.bComms && ingame.side == InGameSide::HOST_OR_SINGLEPLAYER)", "src/titleui/widgets/multilobbyoptions.cpp:502", "initialize", "Not the host?": Exit Function
        widg.setString ""
        If NETGameIsLocked() Then
            WzLobbyPasswordOptionPopover_resetGamePassword strongPopover
        End If

    Case FN_MLO_PW_CHECKBOX_CLICK               ' [weakPopover](W_BUTTON&)
        Set strongPopover = WeakWidget_lock(f.cap(0))
        If Not (Not strongPopover Is Nothing) Then ASSERT_FAILURE "strongPopover != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:519", "initialize", "No parent?": Exit Function
        If Not ((NetPlay_bComms And ingame.side = InGameSide_HOST_OR_SINGLEPLAYER)) Then ASSERT_FAILURE "(NetPlay.bComms && ingame.side == InGameSide::HOST_OR_SINGLEPLAYER)", "src/titleui/widgets/multilobbyoptions.cpp:520", "initialize", "Not the host?": Exit Function
        Set pd = WzLobbyPasswordOptionPopover_(strongPopover)
        If WzAdvCheckbox_isChecked(pd.requiredPasswordCheckbox) Then
            WzLobbyPasswordOptionPopover_setGamePassword strongPopover, pd.passwordEditBox.getString()
        Else
            WzLobbyPasswordOptionPopover_resetGamePassword strongPopover
        End If

    Case FN_displayWrappedMainRoomItem
        displayWrappedMainRoomItem args(0), args(1), args(2)
    Case FN_displayWrappedOptionSectionItem
        displayWrappedOptionSectionItem args(0), args(1), args(2)
    Case FN_displayInlineEditBox
        displayInlineEditBox args(0), args(1), args(2)
    Case FN_gameNameEditHandlerFunc             ' (W_EDITBOX& widg)
        gameNameEditHandlerFunc args(0)

    Case FN_MLO_OPTION_IS_HOST                  ' [](MultibuttonWidget&, int newValue) -> bool
        multilobbyoptions_c_fn = (ingame.side = InGameSide_HOST_OR_SINGLEPLAYER)

    Case FN_MLO_GAMENAME_ESCAPE                 ' [](W_EDITBOX& widg)
        Set widg = args(0)
        widg.setString game.name_

    Case FN_MLO_ROOMNAME_EDIT_CLICK             ' [psWeakGameNameEditBox](W_BUTTON&)
        widgScheduleTask WzFunc_new(FN_MLO_ROOMNAME_EDIT_TASK, f.cap(0))

    Case FN_MLO_ROOMNAME_EDIT_TASK              ' [psWeakGameNameEditBox]
        Set strongEditBox = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongEditBox Is Nothing) Then ASSERT_FAILURE "strongEditBox != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:904", "initialize", "Null edit box": Exit Function
        ' give the edit box the focus
        ' (W_CONTEXT context = W_CONTEXT::ZeroContext(): a new W_CONTEXT is zeroed)
        context.mx = strongEditBox.screenPosX() + strongEditBox.width()
        context.my = strongEditBox.screenPosY()
        W_EDITBOX_simulateClick strongEditBox, context, True

    Case FN_MLO_PASSWORD_BUTTON_CLICK           ' [psWeakMultiLobbyOptions](W_BUTTON&)
        widgScheduleTask WzFunc_new(FN_MLO_PASSWORD_BUTTON_TASK, f.cap(0))

    Case FN_MLO_PASSWORD_BUTTON_TASK            ' [psWeakMultiLobbyOptions]
        Set strongMultiLobbyOptions = WeakWidget_lock(f.cap(0))
        If Not (Not strongMultiLobbyOptions Is Nothing) Then ASSERT_FAILURE "strongMultiLobbyOptions != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:929", "initialize", "No parent": Exit Function
        WzMultiLobbyOptionsImpl_displayPasswordOverlayForm strongMultiLobbyOptions, WzMultiLobbyOptionsImpl_(strongMultiLobbyOptions).roomPasswordOptionsButton

    Case FN_MLO_MAPVIEW_CLICK                   ' [](W_BUTTON&)
        widgScheduleTask WzFunc_new(FN_MLO_MAPVIEW_TASK)

    Case FN_MLO_MAPVIEW_TASK                    ' []
        loadMapPreview True

    Case FN_MLO_MAPCHANGE_CLICK                 ' [psWeakTitleUI](W_BUTTON&)
        Set multiOptionsTitleUI = titleUIFromPtr(f.cap(0))
        If Not (Not multiOptionsTitleUI Is Nothing) Then ASSERT_FAILURE "multiOptionsTitleUI != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:957", "initialize", "Null title UI?": Exit Function
        multiOptionsTitleUI.openMapChooser ' FUTURE TODO: Replace this with a new map chooser overlay screen

    Case FN_MLO_MATCHOPTIONS_CLICK              ' [psWeakMultiLobbyOptions](W_BUTTON&)
        widgScheduleTask WzFunc_new(FN_MLO_MATCHOPTIONS_TASK, f.cap(0))

    Case FN_MLO_MATCHOPTIONS_TASK               ' [psWeakMultiLobbyOptions]
        Set strongMultiLobbyOptions = WeakWidget_lock(f.cap(0))
        If Not (Not strongMultiLobbyOptions Is Nothing) Then ASSERT_FAILURE "strongMultiLobbyOptions != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:978", "initialize", "No parent": Exit Function
        WzMultiLobbyOptionsImpl_displayMatchOptionsMenu strongMultiLobbyOptions, WzMultiLobbyOptionsImpl_(strongMultiLobbyOptions).matchOptionsConfigureButton

    Case FN_MLO_SCAV_CANCHOOSE                  ' [](MultibuttonWidget&, int newValue) -> bool
        multilobbyoptions_c_fn = Not getLockedOptions().scavengers And (ingame.side = InGameSide_HOST_OR_SINGLEPLAYER)

    Case FN_MLO_SCAV_ONCHOOSE                   ' [](MultibuttonWidget& widget, int newValue)
        newValue = args(1)
        If ingame.side = InGameSide_HOST_OR_SINGLEPLAYER Then
            game.scavengers = newValue And &HFF&
            resetReadyStatus False
            If NetPlay_isHost Then
                sendOptions
            End If
        End If

    Case FN_MLO_ALLIANCE_CANCHOOSE
        multilobbyoptions_c_fn = Not getLockedOptions().alliances And (ingame.side = InGameSide_HOST_OR_SINGLEPLAYER)

    Case FN_MLO_ALLIANCE_ONCHOOSE
        newValue = args(1)
        If ingame.side = InGameSide_HOST_OR_SINGLEPLAYER Then
            game.alliance = newValue And &HFF&

            resetReadyStatus False
            netPlayersUpdated = True

            If NetPlay_isHost Then
                sendOptions

                NETsetLobbyConfigFlagsFields game.alliance, game.techLevel, game.power, game.base
            End If
        End If

    Case FN_MLO_POWER_CANCHOOSE
        multilobbyoptions_c_fn = Not getLockedOptions().power And (ingame.side = InGameSide_HOST_OR_SINGLEPLAYER)

    Case FN_MLO_POWER_ONCHOOSE
        newValue = args(1)
        If ingame.side = InGameSide_HOST_OR_SINGLEPLAYER Then
            game.power = newValue
            resetReadyStatus False

            If NetPlay_isHost Then
                sendOptions

                NETsetLobbyConfigFlagsFields game.alliance, game.techLevel, game.power, game.base
            End If
        End If

    Case FN_MLO_BASE_CANCHOOSE
        multilobbyoptions_c_fn = Not getLockedOptions().bases And (ingame.side = InGameSide_HOST_OR_SINGLEPLAYER)

    Case FN_MLO_BASE_ONCHOOSE
        newValue = args(1)
        If ingame.side = InGameSide_HOST_OR_SINGLEPLAYER Then
            game.base = newValue And &HFF&

            resetReadyStatus False

            If NetPlay_isHost Then
                sendOptions

                NETsetLobbyConfigFlagsFields game.alliance, game.techLevel, game.power, game.base
            End If
        End If

    Case FN_MLO_TECH_ONCHOOSE
        newValue = args(1)
        If ingame.side = InGameSide_HOST_OR_SINGLEPLAYER Then
            game.techLevel = newValue

            resetReadyStatus False

            If NetPlay_isHost Then
                sendOptions

                NETsetLobbyConfigFlagsFields game.alliance, game.techLevel, game.power, game.base
            End If
        End If

    Case FN_MLO_LIMITSINFO_CLICK                ' [](W_BUTTON&)
        widgScheduleTask WzFunc_new(FN_MLO_LIMITSINFO_TASK)

    Case FN_MLO_LIMITSINFO_TASK                 ' []()
        printStructureLimitsInfo ingame.structureLimits, ingame.structureLimitsCount, WzFunc_new(FN_MLO_LIMITSINFO_PRINTLINE)

    Case FN_MLO_LIMITSINFO_PRINTLINE            ' [](const std::string& line)
        line_ = args(0)
        addConsoleMessage line_, DEFAULT_JUSTIFY, SYSTEM_MESSAGE

    Case FN_MLO_LIMITSOPTIONS_CLICK             ' [](W_BUTTON&)
        ' POSSIBLE FUTURE TODO: Open context menu?
        widgScheduleTask WzFunc_new(FN_MLO_LIMITSOPTIONS_TASK)

    Case FN_MLO_LIMITSOPTIONS_TASK              ' []()
        If TypeOf wzTitleUICurrent Is WzMultiplayerOptionsTitleUI Then Set parentTitleUI = wzTitleUICurrent
        changeTitleUI WzMultiLimitTitleUI_new(parentTitleUI)

    Case FN_MLO_OPTBUTTON_UPDATE                ' [pref, weakSelf]() (cap(0): T)
        Set strongSelf = WIDGET_fromPtr(f.cap(2))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:661", "addMultiOptButtonT", "No longer valid": Exit Function
        showLocalPrefValue = (ingame.side = InGameSide_MULTIPLAYER_CLIENT) And Not (NetPlay_isHost)
        If showLocalPrefValue Then
            prefValue = getMultiOptionPrefValueT(f.cap(0), f.cap(1))
            WzLobbyMultiOptionButton_setLocalVoteIndicator strongSelf, prefValue
        Else
            WzLobbyMultiOptionButton_setLocalVoteIndicator strongSelf, False
        End If
        showTotalPrefValues = (ingame.side = InGameSide_HOST_OR_SINGLEPLAYER) And NetPlay_isHost And NetPlay_bComms
        If showTotalPrefValues Then
            prefValueTotal = getMultiOptionPrefValueTotalT(f.cap(0), f.cap(1))
            WzLobbyMultiOptionButton_setVoteTotal strongSelf, prefValueTotal
        Else
            WzLobbyMultiOptionButton_setVoteTotal strongSelf, 0    ' setVoteTotal(false)
        End If

    Case FN_MLO_OPTBUTTON_CLICK                 ' [pref](W_BUTTON& b) (cap(0): T)
        Set b = args(0)
        Set psButton = b                        ' std::dynamic_pointer_cast<WzLobbyMultiOptionButton>(b.shared_from_this())
        If Not WIDGET_isKindOf(psButton, K_WzLobbyMultiOptionButton) Then Set psButton = Nothing
        If Not (Not psButton Is Nothing) Then ASSERT_FAILURE "psButton != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:689", "addMultiOptButtonT", "Null button?": Exit Function
        ' Toggle pref value state and update setLocalVoteIndicator
        newPref = Not getMultiOptionPrefValueT(f.cap(0), f.cap(1))
        setMultiOptionPrefValueT f.cap(0), f.cap(1), newPref
        WzLobbyMultiOptionButton_setLocalVoteIndicator psButton, newPref

    Case FN_MLO_PWOVERLAY_CLOSE                 ' [weakPasswordPopover, weakMultiLobbyOptions]()
        ' close handler
        Set strongPasswordPopover = WeakWidget_lock(f.cap(0))
        If Not (Not strongPasswordPopover Is Nothing) Then ASSERT_FAILURE "strongPasswordPopover != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:1344", "displayPasswordOverlayForm", "No widget": Exit Function
        Set strongMultiLobbyOptions = WeakWidget_lock(f.cap(1))
        If Not strongMultiLobbyOptions Is Nothing Then
            WzMultiLobbyOptionsImpl_(strongMultiLobbyOptions).cachedPasswordFieldValue = WzLobbyPasswordOptionPopover_getCurrentPasswordFieldValue(strongPasswordPopover)
            WzMultiLobbyOptionsImpl_(strongMultiLobbyOptions).roomPasswordOptionsButton.setState 0 ' clear clicklock state
        End If

    Case FN_MLO_PWOVERLAY_FOCUS_TASK            ' [weakPasswordPopover]()
        Set strongPasswordPopover = WeakWidget_lock(f.cap(0))
        If Not (Not strongPasswordPopover Is Nothing) Then ASSERT_FAILURE "strongPasswordPopover != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:1354", "displayPasswordOverlayForm", "No widget": Exit Function
        WzLobbyPasswordOptionPopover_givePasswordFieldFocus strongPasswordPopover

    Case FN_MLO_MENUITEM_CLICK                  ' [onClick](WzClickableOptionsChoiceWidget&, WIDGET_KEY key)
        WzFunc_call f.cap(0), args(1)

    Case FN_MLO_MENU_RANDOMIZE                  ' [](WIDGET_KEY)
        multiLobbyRandomizeOptions

    Case FN_MLO_MENU_STARALL                    ' [](WIDGET_KEY)
        If (ingame.side = InGameSide_MULTIPLAYER_CLIENT) And Not (NetPlay_isHost) Then
            setMultiOptionBuiltinPrefValues True
        End If

    Case FN_MLO_MENU_CLEARALL                   ' [](WIDGET_KEY)
        If (ingame.side = InGameSide_MULTIPLAYER_CLIENT) And Not (NetPlay_isHost) Then
            resetMultiOptionPrefValues selectedPlayer
        End If

    Case FN_MLO_MENU_SHOWPREFS                  ' [weakMultiLobbyOptions](WIDGET_KEY) (cap(1): true for "Show", false for "Hide")
        Set strongMultiLobbyOptions = WeakWidget_lock(f.cap(0))
        If f.cap(1) Then
            If Not (Not strongMultiLobbyOptions Is Nothing) Then ASSERT_FAILURE "strongMultiLobbyOptions != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:1417", "buildMatchOptionsMenu", "Missing pointer": Exit Function
        Else
            If Not (Not strongMultiLobbyOptions Is Nothing) Then ASSERT_FAILURE "strongMultiLobbyOptions != nullptr", "src/titleui/widgets/multilobbyoptions.cpp:1425", "buildMatchOptionsMenu", "Missing pointer": Exit Function
        End If
        WzMultiLobbyOptionsImpl_forceSetShowTotalOptionVotes strongMultiLobbyOptions, f.cap(1)

    Case FN_MLO_MATCHMENU_CLOSE                 ' [weakMultiLobbyOptions]()
        ' close handler
        Set strongMultiLobbyOptions = WeakWidget_lock(f.cap(0))
        If Not strongMultiLobbyOptions Is Nothing Then
            WzMultiLobbyOptionsImpl_refreshData strongMultiLobbyOptions
            WzMultiLobbyOptionsImpl_(strongMultiLobbyOptions).matchOptionsConfigureButton.setState 0 ' clear clicklock state
        End If
    End Select
End Function

' (VB) std::weak_ptr<WzMultiplayerOptionsTitleUI>::lock(): the title UI at the address (0: nullptr)
Private Function titleUIFromPtr(ByVal p As Long) As WzMultiplayerOptionsTitleUI
    Dim o As WzMultiplayerOptionsTitleUI
    If p = 0 Then Exit Function
    CopyMemory o, p, 4
    Set titleUIFromPtr = o
    CopyMemory o, 0&, 4
End Function
