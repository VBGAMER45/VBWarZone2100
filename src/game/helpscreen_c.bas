Attribute VB_Name = "helpscreen_c"
' src/screens/helpscreen.cpp, helpscreen.h, whole: the help overlay screen ("help" in the chat box: hci's
' intShowWidgetHelp): a darkening click form over the game's screen with a trigger widget over every visible widget
' that has help or a tool tip; hovering or clicking a trigger shows the widget's help panel. Its widget classes'
' fields are the data classes named as the classes. (VB) struct W_HELP_OVERLAY_SCREEN : public W_SCREEN is a W_SCREEN
' (W_SCREEN_make with its root form, a W_HELPSCREEN_CLICKFORM: rootHelpScreenForm is the screen's psForm) whose
' OnCloseFunc this module keeps ("p<ObjPtr>" in helpScreens, dropped when the screen closes: a closed screen's weak
' pointer locks to nothing, as hci releases it then). The click form and the trigger widgets register in widgweak_vb
' (their weak pointers are serials). HelpTriggerWidget's target is held strongly: the click form's
' registeredWidgets holds every target while its trigger exists (C++'s weak_ptr never expires there), and an address
' could outlive the widget. W_HELPSCREEN_CLICKFORM's registeredWidgets (an unordered_map of shared_ptr<WIDGET> ->
' WidgetInfo{help, helpTriggerWidget}) is a Collection keyed "p<ObjPtr(widget)>" of Array(widget, help, trigger);
' WidgetHelp is copied as C++ copies it (WidgetHelp.copy).
' Translations (_()) are not ported: the English text is used.
' (VB) widget class HelpOverlayCloseButton : W_BUTTON
' (VB) widget class HelpDisplayForm : W_FORM
' (VB) widget class HelpTriggerWidget : WIDGET
' (VB) widget class W_HELPSCREEN_CLICKFORM : W_CLICKFORM
Option Explicit

' the lambdas (WzFunc: helpscreen_c_fn)
Public Const FN_helpOverlay_rootClicked As Long = 18620       ' newRootFrm->onClickedFunc = [psWeakHelpOverlayScreen]() {...}
Public Const FN_helpCloseButton_calcLayout As Long = 18621    ' button->setCalcLayout([minButtonWidthForText](WIDGET *psWidget) {...})
Public Const FN_helpCloseButton_onClick As Long = 18622       ' button->addOnClickHandler([psWeakHelpOverlayScreen](W_BUTTON& button) {...})
Public Const FN_helpTrigger_showTask As Long = 18623          ' widgScheduleTask([triggerWidget]{ triggerWidget->showWidgetHelp(); })
Public Const FN_helpTrigger_closeTask As Long = 18624         ' widgScheduleTask([triggerWidget]{ triggerWidget->closeWidgetHelp(); })
Public Const FN_helpDisplayForm_calcLayout As Long = 18625    ' widgetHelpDisplayForm->setCalcLayout([weakParent](WIDGET *psWidget) {...})
Public Const FN_helpClickForm_calcLayout As Long = 18626      ' W_HELPSCREEN_CLICKFORM::make, sInit.calcLayout

Private helpScreens As Collection       ' (VB) the W_HELP_OVERLAY_SCREENs' OnCloseFunc: "p<ObjPtr>" -> WzFunc (or Nothing)

' MARK: - Custom Close Help UI button

Private Const CloseHelpPadding As Long = 12
Private Const CloseHelpMargin As Long = 5

' class HelpDisplayForm
Private Const INTERNAL_PADDING As Long = 15
Private Const MaxWidth As Long = 450

'--------------------------------------------------------------------------
' (VB) helpers

' const PIELIGHT WZCOL_HELP_UI_BACKGROUND = pal_RGBA(5,9,235,190);
Private Function WZCOL_HELP_UI_BACKGROUND() As Long
    WZCOL_HELP_UI_BACKGROUND = pal_RGBA(5, 9, 235, 190)
End Function

' const PIELIGHT WZCOL_HELP_UI_BACKGROUND_DARK = pal_RGBA(2,5,173,255);
Private Function WZCOL_HELP_UI_BACKGROUND_DARK() As Long
    WZCOL_HELP_UI_BACKGROUND_DARK = pal_RGBA(2, 5, 173, 255)
End Function

' const PIELIGHT WZCOL_HELP_UI_BORDER_LIGHT = pal_RGBA(255, 255, 255, 80);
Private Function WZCOL_HELP_UI_BORDER_LIGHT() As Long
    WZCOL_HELP_UI_BORDER_LIGHT = pal_RGBA(255, 255, 255, 80)
End Function

' a PIELIGHT's alpha (its byte.a: the top byte)
Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

Private Function helpScreens_has(ByVal p As Long) As Boolean
    Dim v As Object
    If helpScreens Is Nothing Or p = 0 Then Exit Function
    On Error GoTo missing
    Set v = helpScreens.Item("p" & p)    ' (VB) Set: the items are WzFuncs or Nothing
    helpScreens_has = True
    Exit Function
missing:
    helpScreens_has = False
End Function

' std::weak_ptr<W_HELP_OVERLAY_SCREEN>::lock(): the screen while it is open, else Nothing
Private Function helpScreen_lock(ByVal p As Long) As W_SCREEN
    If helpScreens_has(p) Then Set helpScreen_lock = W_SCREEN_fromPtr(p)
End Function

' registeredWidgets.find(widget): the entry (Array(widget, help, trigger)), or Empty (end())
Private Function registeredWidgets_find(ByVal self As WIDGET, ByVal widget As WIDGET) As Variant
    On Error GoTo missing
    registeredWidgets_find = W_HELPSCREEN_CLICKFORM_(self).registeredWidgets.Item("p" & ObjPtr(widget))
    Exit Function
missing:
    registeredWidgets_find = Empty
End Function

'--------------------------------------------------------------------------
' MARK: - Custom Close Help UI button

' HelpOverlayCloseButton() : W_BUTTON() {}
Public Sub HelpOverlayCloseButton_construct(ByVal self As WIDGET)
    Dim d As HelpOverlayCloseButton
    W_BUTTON_construct self
    self.kind = K_HelpOverlayCloseButton ' (VB) virtual calls from here on reach HelpOverlayCloseButton's overrides
    Set d = HelpOverlayCloseButton_(self)
    d.downStateMask = WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK
    d.greyStateMask = WBUT_DISABLE
    Set d.text = New WzText
End Sub

' std::make_shared<HelpOverlayCloseButton>()
Private Function HelpOverlayCloseButton_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_HelpOverlayCloseButton)
    HelpOverlayCloseButton_construct widget
    Set HelpOverlayCloseButton_new = widget
End Function

Public Sub HelpOverlayCloseButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As HelpOverlayCloseButton, x0 As Long, y0 As Long, x1 As Long, y1 As Long, haveText As Boolean
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, light_border As Long, fill_color As Long
    Dim c1 As Long, c2 As Long, fw As Long, fx As Long, fy As Long, textColor As Long
    Set d = HelpOverlayCloseButton_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    haveText = Len(W_BUTTON_(self).pText) <> 0

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    ' Display the button.
    light_border = WZCOL_HELP_UI_BORDER_LIGHT()
    If isDown Or isDisabled Then fill_color = WZCOL_HELP_UI_BACKGROUND_DARK() Else fill_color = WZCOL_HELP_UI_BACKGROUND()
    If isDown Then c1 = pal_RGBA(0, 0, 0, 0) Else c1 = light_border
    If isDisabled Then c2 = light_border Else c2 = psPalette(WZCOL_FORM_DARK)
    iV_ShadowBox x0, y0, x1, y1, 0, c1, c2, fill_color
    If isHighlight Then
        iV_Box x0 + 2, y0 + 2, x1 - 2, y1 - 2, psPalette(WZCOL_FORM_HILITE)
    End If

    If haveText Then
        d.text.setText W_BUTTON_(self).pText, W_BUTTON_(self).FontID
        fw = d.text.width()
        fx = x0 + (self.width() - fw) \ 2
        fy = y0 + (self.height() - d.text.lineSize()) \ 2 - d.text.aboveBase()
        textColor = psPalette(WZCOL_TEXT_BRIGHT)
        If isDisabled Then
            textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
        End If
        d.text.render fx, fy, textColor
    End If
End Sub

'--------------------------------------------------------------------------
' MARK: - Help overlay screen

' std::shared_ptr<W_HELP_OVERLAY_SCREEN> W_HELP_OVERLAY_SCREEN::make(const OnCloseFunc& _onCloseFunc)
Public Function W_HELP_OVERLAY_SCREEN_make(ByVal onCloseFunc As WzFunc) As W_SCREEN
    Dim newRootFrm As WIDGET, screen As W_SCREEN, psWeakHelpOverlayScreen As Long, button As WIDGET
    Dim minButtonWidthForText As Long
    ' (class make_shared_enabler: public W_HELP_OVERLAY_SCREEN {};)
    Set newRootFrm = W_HELPSCREEN_CLICKFORM_make()
    Set screen = W_SCREEN_make(newRootFrm)  ' screen->initialize(newRootFrm);
    ' screen->onCloseFunc = _onCloseFunc; (screen->rootHelpScreenForm = newRootFrm: the screen's psForm)
    If helpScreens Is Nothing Then Set helpScreens = New Collection
    If helpScreens_has(ObjPtr(screen)) Then helpScreens.Remove "p" & ObjPtr(screen)  ' (a dead screen's address)
    helpScreens.Add onCloseFunc, "p" & ObjPtr(screen)

    psWeakHelpOverlayScreen = ObjPtr(screen)
    Set W_HELPSCREEN_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_helpOverlay_rootClicked, psWeakHelpOverlayScreen)
    Set W_HELPSCREEN_CLICKFORM_(newRootFrm).onCancelPressed = W_HELPSCREEN_CLICKFORM_(newRootFrm).onClickedFunc

    ' Add "Close Help Mode" button
    Set button = HelpOverlayCloseButton_new()
    button.setString "Close Help Mode"
    W_BUTTON_(button).FontID = font_regular
    minButtonWidthForText = iV_GetTextWidth(W_BUTTON_(button).pText, W_BUTTON_(button).FontID)
    button.setCalcLayout WzFunc_new(FN_helpCloseButton_calcLayout, minButtonWidthForText)
    W_BUTTON_addOnClickHandler button, WzFunc_new(FN_helpCloseButton_onClick, psWeakHelpOverlayScreen)
    newRootFrm.attach button

    Set W_HELP_OVERLAY_SCREEN_make = screen
End Function

Private Sub W_HELP_OVERLAY_SCREEN_closeHelpOverlayScreen(ByVal self As W_SCREEN)
    Dim onCloseFunc As WzFunc, p As Long
    widgRemoveOverlayScreen self
    p = ObjPtr(self)
    If helpScreens_has(p) Then
        Set onCloseFunc = helpScreens.Item("p" & p)
        helpScreens.Remove "p" & p      ' (VB) the screen is closed: its weak pointers lock to nothing from now on
    End If
    If Not onCloseFunc Is Nothing Then
        WzFunc_call onCloseFunc, self
    End If
End Sub

Public Sub W_HELP_OVERLAY_SCREEN_setHelpFromWidgets(ByVal screen As W_SCREEN, ByVal root As WIDGET)
    Dim i As Long, child As WIDGET, pHelp As WidgetHelp, childTipStr As String, firstLinePos As Long
    Dim generatedHelp As WidgetHelp
    If Not (Not root Is Nothing) Then ASSERT_FAILURE "root != nullptr", "src/screens/helpscreen.cpp:154", "setHelpFromWidgets", "Null widget?": Exit Sub
    ' (Enumerate the child widgets in decreasing z-order (i.e. "top-down"))
    For i = root.childCount() - 1 To 0 Step -1
        Set child = root.childAt(i)
        If Not child.visible() Then
            GoTo continue_
        End If
        W_HELP_OVERLAY_SCREEN_setHelpFromWidgets screen, child
        ' Use the widget's help, if available
        Set pHelp = child.getHelp()
        If Not pHelp Is Nothing Then
            W_HELP_OVERLAY_SCREEN_setHelpForWidget screen, child, pHelp
        Else
            ' Try the widget tooltip, if available
            childTipStr = child.getTip()
            If Len(childTipStr) = 0 Then
                GoTo continue_
            End If
            ' Parse the tip
            ' - First line as title
            ' - Additional lines as description
            firstLinePos = InStr(childTipStr, vbLf)
            If firstLinePos <> 0 Then
                Set generatedHelp = New WidgetHelp
                generatedHelp.setTitle Left$(childTipStr, firstLinePos - 1)
                If firstLinePos < Len(childTipStr) Then
                    generatedHelp.setDescription Mid$(childTipStr, firstLinePos + 1)
                End If
                W_HELP_OVERLAY_SCREEN_setHelpForWidget screen, child, generatedHelp
            Else
                Set generatedHelp = New WidgetHelp
                W_HELP_OVERLAY_SCREEN_setHelpForWidget screen, child, generatedHelp.setTitle(childTipStr)
            End If
        End If
continue_:
    Next
End Sub

Private Sub W_HELP_OVERLAY_SCREEN_setHelpForWidget(ByVal self As W_SCREEN, ByVal widget As WIDGET, ByVal help As WidgetHelp)
    W_HELPSCREEN_CLICKFORM_setHelpForWidget self.psForm, widget, help
End Sub

' (no caller in 4.7.0)
Private Sub W_HELP_OVERLAY_SCREEN_removeHelpForWidget(ByVal self As W_SCREEN, ByVal widget As WIDGET)
    W_HELPSCREEN_CLICKFORM_removeHelpForWidget self.psForm, widget
End Sub

'--------------------------------------------------------------------------
' MARK: - HelpDisplayForm

' std::string to_string(const WidgetHelp::InteractionTriggers& trigger)
Private Function to_string_InteractionTrigger(ByVal trigger As InteractionTriggers) As String
    Select Case trigger
    Case InteractionTriggers_PrimaryClick
        to_string_InteractionTrigger = "Click/Tap"
        Exit Function
    Case InteractionTriggers_SecondaryClick
        to_string_InteractionTrigger = "Right-Click"
        Exit Function
    Case InteractionTriggers_ClickAndHold
        to_string_InteractionTrigger = "Click/Tap and Hold"
        Exit Function
    Case InteractionTriggers_Misc
        to_string_InteractionTrigger = "Other"
        Exit Function
    End Select
    to_string_InteractionTrigger = "" ' silence warning
End Function

' std::string to_string(const std::set<WidgetHelp::InteractionTriggers>& triggers): the set as its mask (its
' triggers in ascending order)
Private Function to_string_InteractionTriggers(ByVal triggers As Long) As String
    Dim result As String, trigger As Long
    result = ""
    For trigger = InteractionTriggers_PrimaryClick To InteractionTriggers_Misc
        If (triggers And CLng(2 ^ trigger)) <> 0 Then
            If Len(result) <> 0 Then
                result = result & vbLf
            End If
            result = result & to_string_InteractionTrigger(trigger)
        End If
    Next
    to_string_InteractionTriggers = result
End Function

' HelpDisplayForm() { }
Public Sub HelpDisplayForm_construct(ByVal self As WIDGET)
    W_FORM_construct self
    self.kind = K_HelpDisplayForm ' (VB) virtual calls from here on reach HelpDisplayForm's overrides
End Sub

Private Sub HelpDisplayForm_initialize(ByVal self As WIDGET, ByVal help As WidgetHelp)
    Dim d As HelpDisplayForm, grid As WIDGET, row As grid_allocation_slot, interactions_label As WIDGET, i As Long
    Dim interaction_type As WIDGET, interaction_effect As WIDGET, col As grid_allocation_slot
    Set d = HelpDisplayForm_(self)
    ' Create title (bold)
    Set d.title_ = W_LABEL_new()
    W_LABEL_setFont d.title_, font_regular_bold, psPalette(WZCOL_TEXT_BRIGHT)
    d.title_.setString help.title_
    self.attach d.title_

    ' Create description (paragraph)
    If Len(help.description) <> 0 Then
        Set d.description = Paragraph_new()
        Paragraph_setFontColour d.description, psPalette(WZCOL_TEXT_BRIGHT)
        Paragraph_setLineSpacing d.description, 2
        Paragraph_setFont d.description, font_regular
        Paragraph_addText d.description, help.description
        self.attach d.description
    End If

    ' Create a list of interactions
    If help.interactionCount() <> 0 Then
        Set grid = GridLayout_new()
        row = grid_allocation_slot_(0)

        Set interactions_label = W_LABEL_new()
        W_LABEL_setFont interactions_label, font_regular, psPalette(WZCOL_TEXT_MEDIUM)
        interactions_label.setString "Interactions:"
        interactions_label.setGeometry 0, 0, W_LABEL_getMaxLineWidth(interactions_label), W_LABEL_requiredHeight(interactions_label)
        col = grid_allocation_slot_(0, 2)
        GridLayout_place grid, col, row, Margin_wrap(5, 0, 5, 0, interactions_label)
        row.start = row.start + 1

        For i = 0 To help.interactionCount() - 1
            Set interaction_type = Paragraph_new()
            Paragraph_setFontColour interaction_type, psPalette(WZCOL_TEXT_BRIGHT)
            Paragraph_setLineSpacing interaction_type, 0
            Paragraph_setFont interaction_type, font_regular_bold
            Paragraph_addText interaction_type, to_string_InteractionTriggers(help.interactionTriggers(i))
            interaction_type.setGeometry 0, 0, (MaxWidth \ 2) - (5 * 2), 0

            Set interaction_effect = Paragraph_new()
            Paragraph_setFontColour interaction_effect, psPalette(WZCOL_TEXT_BRIGHT)
            Paragraph_setLineSpacing interaction_effect, 0
            Paragraph_setFont interaction_effect, font_regular
            Paragraph_addText interaction_effect, help.interactionDescription(i)
            interaction_effect.setGeometry 0, 0, MaxWidth \ 2, 0

            col = grid_allocation_slot_(0)
            GridLayout_place grid, col, row, Margin_wrap(5, 5, 5, 5, interaction_type)
            col = grid_allocation_slot_(1, 1, False)
            GridLayout_place grid, col, row, Alignment_wrap(Alignment_(Alignment_Vertical_Center, Alignment_Horizontal_Left), Margin_wrap(5, 5, 5, 5, interaction_effect))
            row.start = row.start + 1
        Next

        grid.setGeometry 0, 0, MaxWidth, grid.idealHeight()

        Set d.interactions = ScrollableListWidget_make()
        ScrollableListWidget_addItem d.interactions, grid
        self.attach d.interactions
    End If

    HelpDisplayForm_buildRelatedKeybindingsList self, help

    HelpDisplayForm_recalcLayout self
End Sub

Private Sub HelpDisplayForm_buildRelatedKeybindingsList(ByVal self As WIDGET, ByVal help As WidgetHelp)
    Dim d As HelpDisplayForm, grid As WIDGET, row As grid_allocation_slot, keys_label As WIDGET, i As Long
    Dim primary_mapping As Long, secondary_mapping As Long, keybindingsStr As String, keyDisplayName As String
    Dim interaction_type As WIDGET, interaction_effect As WIDGET, col As grid_allocation_slot, w As Long
    Set d = HelpDisplayForm_(self)
    If help.relatedKeybindingCount() = 0 Then
        Exit Sub
    End If

    Set grid = GridLayout_new()
    row = grid_allocation_slot_(0)

    Set keys_label = W_LABEL_new()
    W_LABEL_setFont keys_label, font_regular, psPalette(WZCOL_TEXT_MEDIUM)
    keys_label.setString "Related Keymappings:"
    keys_label.setGeometry 0, 0, W_LABEL_getMaxLineWidth(keys_label), W_LABEL_requiredHeight(keys_label)
    col = grid_allocation_slot_(0, 2)
    GridLayout_place grid, col, row, Margin_wrap(5, 0, 5, 0, keys_label)
    row.start = row.start + 1

    For i = 0 To help.relatedKeybindingCount() - 1
        primary_mapping = gInputManager_mappings_get_name(help.relatedKeybinding(i), KeyMappingSlot_PRIMARY)
        secondary_mapping = gInputManager_mappings_get_name(help.relatedKeybinding(i), KeyMappingSlot_SECONDARY)
        keybindingsStr = ""
        keyDisplayName = ""
        If primary_mapping <> 0 Then
            If Len(keybindingsStr) <> 0 Then keybindingsStr = keybindingsStr & vbLf
            keybindingsStr = keybindingsStr & KeyMapping_toString(primary_mapping)
            keyDisplayName = KeyFunctionInfo_displayName(KeyMapping_info(primary_mapping))
        End If
        If secondary_mapping <> 0 Then
            If Len(keybindingsStr) <> 0 Then keybindingsStr = keybindingsStr & vbLf
            keybindingsStr = keybindingsStr & KeyMapping_toString(secondary_mapping)
            If Len(keyDisplayName) = 0 Then
                keyDisplayName = KeyFunctionInfo_displayName(KeyMapping_info(secondary_mapping))
            End If
        End If
        If Len(keybindingsStr) = 0 Or Len(keyDisplayName) = 0 Then
            GoTo continue_
        End If
        Set interaction_type = W_LABEL_new()
        W_LABEL_setFont interaction_type, font_regular_bold, psPalette(WZCOL_TEXT_MEDIUM)
        interaction_type.setString keyDisplayName
        w = W_LABEL_getMaxLineWidth(interaction_type)
        If MaxWidth \ 2 < w Then w = MaxWidth \ 2
        interaction_type.setGeometry 0, 0, w, W_LABEL_requiredHeight(interaction_type)

        Set interaction_effect = Paragraph_new()
        Paragraph_setFontColour interaction_effect, psPalette(WZCOL_TEXT_MEDIUM)
        Paragraph_setLineSpacing interaction_effect, 0
        Paragraph_setFont interaction_effect, font_regular
        Paragraph_addText interaction_effect, keybindingsStr
        interaction_effect.setGeometry 0, 0, MaxWidth \ 2, 0

        col = grid_allocation_slot_(0)
        GridLayout_place grid, col, row, Margin_wrap(2, 5, 2, 5, interaction_type)
        col = grid_allocation_slot_(1, 1, False)
        GridLayout_place grid, col, row, Alignment_wrap(Alignment_(Alignment_Vertical_Center, Alignment_Horizontal_Left), Margin_wrap(2, 5, 2, 5, interaction_effect))
        row.start = row.start + 1
continue_:
    Next

    grid.setGeometry 0, 0, MaxWidth, grid.idealHeight()

    Set d.relatedKeybindings = ScrollableListWidget_make()
    ScrollableListWidget_addItem d.relatedKeybindings, grid
    self.attach d.relatedKeybindings
End Sub

' static std::shared_ptr<HelpDisplayForm> make(const WidgetHelp& help)
Private Function HelpDisplayForm_make(ByVal help As WidgetHelp) As WIDGET
    Dim widget As WIDGET
    ' (class make_shared_enabler: public HelpDisplayForm {};)
    Set widget = WIDGET_alloc(K_HelpDisplayForm)
    HelpDisplayForm_construct widget
    HelpDisplayForm_initialize widget, help
    Set HelpDisplayForm_make = widget
End Function

Public Sub HelpDisplayForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long
    If (self.style And WFORM_INVISIBLE) = 0 Then
        x0 = self.x() + xOffset
        y0 = self.y() + yOffset
        x1 = x0 + self.width()
        y1 = y0 + self.height()

        pie_UniTransBoxFill x0, y0, x1, y1, WZCOL_HELP_UI_BACKGROUND()
        iV_Box x0 + 1, y0 + 1, x1 + 1, y1 + 1, WZCOL_HELP_UI_BACKGROUND_DARK()
    End If
End Sub

Public Sub HelpDisplayForm_geometryChanged(ByVal self As WIDGET)
    ' (VB) the widget constructor's hazard (CONVENTIONS, "Widgets"): W_FORM's construction reaches here before
    ' initialize made the title
    If HelpDisplayForm_(self).title_ Is Nothing Then Exit Sub
    HelpDisplayForm_recalcLayout self
End Sub

Public Function HelpDisplayForm_idealWidth(ByVal self As WIDGET) As Long
    Dim d As HelpDisplayForm
    Set d = HelpDisplayForm_(self)
    If Not d.description Is Nothing Or Not d.interactions Is Nothing Then
        HelpDisplayForm_idealWidth = MaxWidth
    Else
        ' only title
        HelpDisplayForm_idealWidth = W_LABEL_getMaxLineWidth(d.title_) + (INTERNAL_PADDING * 2)
    End If
End Function

Public Function HelpDisplayForm_idealHeight(ByVal self As WIDGET) As Long
    Dim d As HelpDisplayForm, bottomOfLastElement As Long
    Set d = HelpDisplayForm_(self)
    bottomOfLastElement = 0
    If Not d.relatedKeybindings Is Nothing Then
        bottomOfLastElement = WzRect_bottom(d.relatedKeybindings.geometry())
    ElseIf Not d.interactions Is Nothing Then
        bottomOfLastElement = WzRect_bottom(d.interactions.geometry())
    ElseIf Not d.description Is Nothing Then
        bottomOfLastElement = WzRect_bottom(d.description.geometry())
    ElseIf Not d.title_ Is Nothing Then
        bottomOfLastElement = INTERNAL_PADDING + W_LABEL_requiredHeight(d.title_)
    End If
    HelpDisplayForm_idealHeight = bottomOfLastElement + INTERNAL_PADDING
End Function

Private Sub HelpDisplayForm_recalcLayout(ByVal self As WIDGET)
    Dim d As HelpDisplayForm, maxChildWidth As Long, w As Long, bottomOfLastElement As Long
    Dim descriptionY As Long, interactionsY As Long, relatedKeybindingsY As Long
    Const contentsPadding As Long = INTERNAL_PADDING
    Set d = HelpDisplayForm_(self)
    maxChildWidth = self.width() - (INTERNAL_PADDING * 2)

    w = W_LABEL_getMaxLineWidth(d.title_)
    If maxChildWidth < w Then w = maxChildWidth
    d.title_.setGeometry INTERNAL_PADDING, INTERNAL_PADDING, w, W_LABEL_requiredHeight(d.title_)

    bottomOfLastElement = d.title_.y() + d.title_.height()

    If Not d.description Is Nothing Then
        descriptionY = bottomOfLastElement + contentsPadding
        d.description.setGeometry INTERNAL_PADDING, descriptionY, maxChildWidth, d.description.height()
        bottomOfLastElement = d.description.y() + d.description.height()
    End If

    If Not d.interactions Is Nothing Then
        interactionsY = bottomOfLastElement + contentsPadding
        d.interactions.setGeometry INTERNAL_PADDING, interactionsY, maxChildWidth, d.interactions.idealHeight()
        d.interactions.setGeometry INTERNAL_PADDING, interactionsY, maxChildWidth, d.interactions.idealHeight() ' call again using calculated height from width change
        bottomOfLastElement = d.interactions.y() + d.interactions.height()
    End If

    If Not d.relatedKeybindings Is Nothing Then
        relatedKeybindingsY = bottomOfLastElement + contentsPadding
        d.relatedKeybindings.setGeometry INTERNAL_PADDING, relatedKeybindingsY, maxChildWidth, d.relatedKeybindings.idealHeight()
        d.relatedKeybindings.setGeometry INTERNAL_PADDING, relatedKeybindingsY, maxChildWidth, d.relatedKeybindings.idealHeight() ' call again using calculated height from width change
        ' (4.7.0 reads interactions here: null when the help has no interactions; the port skips it then)
        If Not d.interactions Is Nothing Then bottomOfLastElement = d.interactions.y() + d.interactions.height()
    End If
End Sub

'--------------------------------------------------------------------------
' MARK: - HelpTriggerWidget

' HelpTriggerWidget() {}
Public Sub HelpTriggerWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_HelpTriggerWidget ' (VB) virtual calls from here on reach HelpTriggerWidget's overrides
    HelpTriggerWidget_(self).weakSerial = WeakWidget_register(self)
End Sub

' (~HelpTriggerWidget(): the widgweak_vb registration ends)
Public Sub HelpTriggerWidget_destruct(ByVal self As WIDGET)
    WeakWidget_unregister HelpTriggerWidget_(self).weakSerial
End Sub

Private Sub HelpTriggerWidget_initialize(ByVal self As WIDGET, ByVal parent As WIDGET, ByVal target As WIDGET)
    Dim d As HelpTriggerWidget
    Set d = HelpTriggerWidget_(self)
    d.parent = W_HELPSCREEN_CLICKFORM_(parent).weakSerial
    Set d.target = target
End Sub

' static std::shared_ptr<HelpTriggerWidget> make(const std::shared_ptr<W_HELPSCREEN_CLICKFORM>& parent,
' const std::shared_ptr<WIDGET>& target)
Private Function HelpTriggerWidget_make(ByVal parent As WIDGET, ByVal target As WIDGET) As WIDGET
    Dim widget As WIDGET, r As WzRect
    ' (class make_shared_enabler: public HelpTriggerWidget {};)
    Set widget = WIDGET_alloc(K_HelpTriggerWidget)
    HelpTriggerWidget_construct widget
    HelpTriggerWidget_initialize widget, parent, target
    r = HelpTriggerWidget_calculateIdealScreenPos(widget)
    widget.setGeometryFromScreenRect r
    Set HelpTriggerWidget_make = widget
End Function

Public Sub HelpTriggerWidget_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As HelpTriggerWidget, strongParent As WIDGET
    Set d = HelpTriggerWidget_(self)
    Set strongParent = WeakWidget_lock(d.parent)
    If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/screens/helpscreen.cpp:557", "clicked", "Parent no longer exists?": Exit Sub

    d.clickedDownOnThis = True

    If Not W_HELPSCREEN_CLICKFORM_helperWidgetHasFocusLock(strongParent) Then
        ' immediately trigger the widget help popup
        widgScheduleTask WzFunc_new(FN_helpTrigger_showTask, self)
    Else
        d.queuedWidgetShowOnRelease = True
    End If
End Sub

Public Sub HelpTriggerWidget_focusLost(ByVal self As WIDGET)
    Dim d As HelpTriggerWidget
    Set d = HelpTriggerWidget_(self)
    widgScheduleTask WzFunc_new(FN_helpTrigger_closeTask, self)
    d.hasFocus = False
    d.clickedDownOnThis = False
    d.queuedWidgetShowOnRelease = False
End Sub

Public Sub HelpTriggerWidget_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As HelpTriggerWidget, lockedScreen As W_SCREEN
    Set d = HelpTriggerWidget_(self)
    If d.clickedDownOnThis Then
        ' Tell the screen that this has focus
        Set lockedScreen = self.screenPointer
        If Not lockedScreen Is Nothing Then
            lockedScreen.setFocus self
            d.hasFocus = True
        End If

        d.clickedDownOnThis = False

        If d.queuedWidgetShowOnRelease Then
            widgScheduleTask WzFunc_new(FN_helpTrigger_showTask, self)
            d.queuedWidgetShowOnRelease = False
        End If
    End If
End Sub

Public Sub HelpTriggerWidget_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As HelpTriggerWidget
    Set d = HelpTriggerWidget_(self)
    d.has_hoverStart = True
    d.hoverStart = widgSteadyNowMs()        ' std::chrono::steady_clock::now()
End Sub

Public Sub HelpTriggerWidget_highlightLost(ByVal self As WIDGET)
    Dim d As HelpTriggerWidget
    Set d = HelpTriggerWidget_(self)
    d.has_hoverStart = False                ' hoverStart = nullopt;
    d.isHighlight = False
    d.clickedDownOnThis = False
    d.queuedWidgetShowOnRelease = False
    If Not d.hasFocus Then
        widgScheduleTask WzFunc_new(FN_helpTrigger_closeTask, self)
    End If
End Sub

Public Sub HelpTriggerWidget_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As HelpTriggerWidget, strongScreen As W_SCREEN, focusWidget As WIDGET
    Set d = HelpTriggerWidget_(self)
    If d.has_hoverStart Then
        ' require hover to be maintained for a minimum duration before triggering the help for this widget
        If Fix(widgSteadyNowMs() - d.hoverStart) >= 100 Then
            d.has_hoverStart = False
            d.isHighlight = True
            Set strongScreen = self.screenPointer
            If Not strongScreen Is Nothing Then
                Set focusWidget = strongScreen.getWidgetWithFocus()
                If Not focusWidget Is Nothing And Not focusWidget Is strongScreen.psForm Then
                    ' another trigger has the focus lock, due to a click
                    ' do not show widget help
                    Exit Sub
                End If
            End If
            widgScheduleTask WzFunc_new(FN_helpTrigger_showTask, self)
        End If
    End If
End Sub

Private Sub DrawAnimatedIndicator(ByVal x0 As Long, ByVal y0 As Long, ByVal size As Long, ByVal indicatorColor As Long, indicatorAnimationProgress As Single, lastUpdatedIndicatorAnimationTime As Long)
    Dim changeAmount As Single, indicator_animation_size As Single, indicator_animation_offset As Single
    Dim animation_progress_inv_val As Single, indicator_animation_alpha As Long, animatedX0 As Single, animatedY0 As Single
    Dim r As Long, g As Long, b As Long
    If Not ULT(realTime - lastUpdatedIndicatorAnimationTime, 50) Then
        changeAmount = 100! * CSng(U32(realTime - lastUpdatedIndicatorAnimationTime)) / CSng(GAME_TICKS_PER_SEC)
        indicatorAnimationProgress = (indicatorAnimationProgress + changeAmount)
        Do While indicatorAnimationProgress > 100!
            indicatorAnimationProgress = indicatorAnimationProgress - 100!
        Loop
        lastUpdatedIndicatorAnimationTime = realTime
    End If

    indicator_animation_size = CSng(size) + (CSng(indicatorAnimationProgress * 5!) / 100!)
    indicator_animation_offset = (indicator_animation_size - CSng(size)) / 2!
    animation_progress_inv_val = 1! - (indicatorAnimationProgress / 100!)
    indicator_animation_alpha = CLng(Fix(animation_progress_inv_val * 255!)) And &HFF&  ' static_cast<UBYTE>
    r = indicatorColor And &HFF&
    g = SHRU(indicatorColor, 8) And &HFF&
    b = SHRU(indicatorColor, 16) And &HFF&

    If IntImages = 0 Then
        ' Interface Images are not loaded - just use simple rectangles

        animatedX0 = CSng(x0) - indicator_animation_offset
        animatedY0 = CSng(y0) - indicator_animation_offset
        pie_UniTransBoxFill animatedX0, animatedY0, animatedX0 + indicator_animation_size, animatedY0 + indicator_animation_size, pal_RGBA(r, g, b, indicator_animation_alpha)

        pie_UniTransBoxFill x0, y0, x0 + size, y0 + size, indicatorColor
        Exit Sub
    End If

    iV_DrawImageTint IntImages, IMAGE_INDICATOR_DOT, x0, y0, indicatorColor, True, size, size

    iV_DrawImageTint IntImages, IMAGE_INDICATOR_DOT_EXPAND, CSng(x0) - indicator_animation_offset, CSng(y0) - indicator_animation_offset, pal_RGBA(r, g, b, indicator_animation_alpha), True, indicator_animation_size, indicator_animation_size
End Sub

Public Sub HelpTriggerWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As HelpTriggerWidget, x0 As Long, y0 As Long, x1 As Long, y1 As Long, indicator_x0 As Long, indicator_y0 As Long
    Dim indicatorColor As Long, progress As Single, lastTime As Long
    Const indicator_size As Long = 10
    Set d = HelpTriggerWidget_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    If d.isHighlight Or d.hasFocus Then
        iV_Box x0 - 1, y0 - 1, x1 + 1, y1 + 1, pal_RGBA(255, 255, 255, 125)
        Exit Sub
    End If

    indicator_x0 = x1 - (indicator_size + d.indicator_inset_x)
    indicator_y0 = y0 + d.indicator_inset_y

    indicatorColor = pal_RGBA(35, 252, 115, 255)

    progress = d.indicatorAnimationProgress
    lastTime = d.lastUpdatedIndicatorAnimationTime
    DrawAnimatedIndicator indicator_x0, indicator_y0, indicator_size, indicatorColor, progress, lastTime
    d.indicatorAnimationProgress = progress
    d.lastUpdatedIndicatorAnimationTime = lastTime
End Sub

' WzRect calculateIdealScreenPos()
Private Function HelpTriggerWidget_calculateIdealScreenPos(ByVal self As WIDGET) As WzRect
    Dim d As HelpTriggerWidget, strongTarget As WIDGET, strongParent As WIDGET
    Set d = HelpTriggerWidget_(self)
    Set strongTarget = d.target
    Set strongParent = WeakWidget_lock(d.parent)
    If strongTarget Is Nothing Then
        ' target no longer exists? clean up this trigger widget
        self.deleteLater
        Exit Function
    End If
    If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/screens/helpscreen.cpp:724", "calculateIdealScreenPos", "Parent no longer exists?": Exit Function
    HelpTriggerWidget_calculateIdealScreenPos = HelpTriggerWidget_calculateIdealScreenPos2(self, strongParent, strongTarget)
End Function

' WzRect calculateIdealScreenPos(const std::shared_ptr<WIDGET> _parent, const std::shared_ptr<WIDGET>& _target)
Private Function HelpTriggerWidget_calculateIdealScreenPos2(ByVal self As WIDGET, ByVal parent As WIDGET, ByVal target As WIDGET) As WzRect
    Dim targetScreenGeometry As WzRect
    If Not (Not parent Is Nothing) Then ASSERT_FAILURE "_parent != nullptr", "src/screens/helpscreen.cpp:730", "calculateIdealScreenPos", "Parent no longer exists?": Exit Function
    targetScreenGeometry = target.screenGeometry()

    ' Overlay the target's screen geometry, in full
    HelpTriggerWidget_calculateIdealScreenPos2 = targetScreenGeometry
End Function

Public Sub HelpTriggerWidget_screenSizeDidChange(ByVal self As WIDGET, ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    Dim r As WzRect
    ' reposition based on the target widget location (in case that changed)
    r = HelpTriggerWidget_calculateIdealScreenPos(self)
    self.setGeometryFromScreenRect r
End Sub

Public Sub HelpTriggerWidget_setGeometryFromScreenRect(ByVal self As WIDGET, r As WzRect)
    Dim strongParent As WIDGET, xOffset As Long, yOffset As Long
    Set strongParent = WeakWidget_lock(HelpTriggerWidget_(self).parent)
    If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/screens/helpscreen.cpp:746", "setGeometryFromScreenRect", "Parent no longer exists?": Exit Sub
    xOffset = strongParent.screenPosX()
    yOffset = strongParent.screenPosY()
    self.setGeometry r.x - xOffset, r.y - yOffset, r.w, r.h
End Sub

Private Sub HelpTriggerWidget_showWidgetHelp(ByVal self As WIDGET)
    Dim d As HelpTriggerWidget, strongTarget As WIDGET, strongParent As WIDGET, pWidgetHelp As WidgetHelp
    Dim widgetHelpDisplayForm As WIDGET, weakParent As Long
    Set d = HelpTriggerWidget_(self)
    Set strongTarget = d.target
    Set strongParent = WeakWidget_lock(d.parent)
    If strongTarget Is Nothing Then
        ' target no longer exists? clean up this trigger widget
        self.deleteLater
        Exit Sub
    End If
    If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/screens/helpscreen.cpp:763", "showWidgetHelp", "Parent no longer exists?": Exit Sub

    ' Get help from parent
    Set pWidgetHelp = W_HELPSCREEN_CLICKFORM_getHelpForWidget(strongParent, strongTarget)
    If pWidgetHelp Is Nothing Then
        Exit Sub
    End If

    Set widgetHelpDisplayForm = HelpDisplayForm_make(pWidgetHelp)

    ' Figure out tbe actual needed height based on the ideal width
    widgetHelpDisplayForm.setGeometry 100, 100, widgetHelpDisplayForm.idealWidth(), widgetHelpDisplayForm.idealHeight()
    widgetHelpDisplayForm.setGeometry 100, 100, widgetHelpDisplayForm.idealWidth(), widgetHelpDisplayForm.idealHeight() ' twice to pick up calculated idealHeight for this width

    ' Figure out position based on parent trigger widget
    weakParent = d.weakSerial
    widgetHelpDisplayForm.setCalcLayout WzFunc_new(FN_helpDisplayForm_calcLayout, weakParent)

    ' Inform parent (so it can highlight the current interface element)
    W_HELPSCREEN_CLICKFORM_setCurrentHelpWidget strongParent, self, widgetHelpDisplayForm
End Sub

Private Sub HelpTriggerWidget_closeWidgetHelp(ByVal self As WIDGET)
    Dim d As HelpTriggerWidget, strongTarget As WIDGET, strongParent As WIDGET
    Set d = HelpTriggerWidget_(self)
    Set strongTarget = d.target
    Set strongParent = WeakWidget_lock(d.parent)
    If strongTarget Is Nothing Then
        ' target no longer exists? clean up this trigger widget
        self.deleteLater
        Exit Sub
    End If
    If strongParent Is Nothing Then
        Exit Sub
    End If
    W_HELPSCREEN_CLICKFORM_unsetCurrentHelpWidget strongParent, self
End Sub

'--------------------------------------------------------------------------
' MARK: - W_FULLSCREENOVERLAY_CLICKFORM

' W_HELPSCREEN_CLICKFORM(W_FORMINIT const *init) : W_CLICKFORM(init) {}
Public Sub W_HELPSCREEN_CLICKFORM_construct_init(ByVal self As WIDGET, init As W_FORMINIT)
    Dim d As W_HELPSCREEN_CLICKFORM
    W_CLICKFORM_construct_init self, init
    self.kind = K_W_HELPSCREEN_CLICKFORM ' (VB) virtual calls from here on reach W_HELPSCREEN_CLICKFORM's overrides
    Set d = W_HELPSCREEN_CLICKFORM_(self)
    d.backgroundColor = pal_RGBA(0, 0, 0, 125)
    Set d.registeredWidgets = New Collection
    d.weakSerial = WeakWidget_register(self)
End Sub

' (~W_HELPSCREEN_CLICKFORM(): the widgweak_vb registration ends)
Public Sub W_HELPSCREEN_CLICKFORM_destruct(ByVal self As WIDGET)
    WeakWidget_unregister W_HELPSCREEN_CLICKFORM_(self).weakSerial
End Sub

' static std::shared_ptr<W_HELPSCREEN_CLICKFORM> make(UDWORD formID = 0)
Private Function W_HELPSCREEN_CLICKFORM_make(Optional ByVal formID As Long = 0) As WIDGET
    Dim sInit As W_FORMINIT, widget As WIDGET
    W_FORMINIT_construct sInit
    sInit.id = formID
    sInit.style = WFORM_PLAIN Or WFORM_CLICKABLE
    sInit.x = 0
    sInit.y = 0
    sInit.width = screenWidth - 1
    sInit.height = screenHeight - 1
    Set sInit.calcLayout = WzFunc_new(FN_helpClickForm_calcLayout)

    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_W_HELPSCREEN_CLICKFORM)
    W_HELPSCREEN_CLICKFORM_construct_init widget, sInit
    Set W_HELPSCREEN_CLICKFORM_make = widget
End Function

Public Sub W_HELPSCREEN_CLICKFORM_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As W_HELPSCREEN_CLICKFORM, lockedScreen As W_SCREEN, widgetWithFocus As WIDGET
    Set d = W_HELPSCREEN_CLICKFORM_(self)
    Set lockedScreen = self.screenPointer
    If Not lockedScreen Is Nothing Then
        Set widgetWithFocus = lockedScreen.getWidgetWithFocus()
        If Not widgetWithFocus Is Nothing And Not widgetWithFocus Is self Then
            ' something else has focus
            lockedScreen.setFocus Nothing
            Exit Sub
        End If
    End If
    If Not d.onClickedFunc Is Nothing Then
        WzFunc_call d.onClickedFunc
    End If
End Sub

Public Sub W_HELPSCREEN_CLICKFORM_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_HELPSCREEN_CLICKFORM, x0 As Long, y0 As Long, strongCutoutWidget As WIDGET, screenRect As WzRect
    Dim currentHelpWidgetScreenRect As WzRect, checkedParent As WIDGET, cutoutRect As WzRect
    Dim surroundingRects(0 To 3) As WzRect, i As Long
    Set d = W_HELPSCREEN_CLICKFORM_(self)
    If Not self.visible() Then
        Exit Sub ' skip if hidden
    End If

    If pielightAlpha(d.backgroundColor) = 0 Then   ' backgroundColor.isTransparent()
        Exit Sub
    End If

    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    Set strongCutoutWidget = WeakWidget_lock(d.currentHelpWidget)
    screenRect = self.screenGeometry()
    If Not strongCutoutWidget Is Nothing Then
        ' Any widget that has help is expected to have a parent (i.e. actually be on a screen)
        Set checkedParent = strongCutoutWidget.parent()
        If Not checkedParent Is Nothing Then
            currentHelpWidgetScreenRect = strongCutoutWidget.screenGeometry()
        Else
            ' This is probably an old widget that gets deleted and recreated constantly (there are a few of these left, sigh)
            ' Skip trying to do a cutout, since we can't determine an updated screen position (as it has no parent)
            Set strongCutoutWidget = Nothing
        End If
    End If
    If Not strongCutoutWidget Is Nothing Then cutoutRect = WzRect_intersectionWith(screenRect, currentHelpWidgetScreenRect)

    If cutoutRect.w <= 0 Or cutoutRect.h <= 0 Then
        ' simple path - draw background over everything
        pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), d.backgroundColor
    Else
        ' left column until left edge of cutout rect
        surroundingRects(0) = WzRect_corners(WzRect_left(screenRect), WzRect_top(screenRect), WzRect_left(cutoutRect), WzRect_bottom(screenRect))
        ' top band above cutout rect
        surroundingRects(1) = WzRect_corners(WzRect_left(cutoutRect), WzRect_top(screenRect), WzRect_right(cutoutRect), WzRect_top(cutoutRect))
        ' right column after cutout rect
        surroundingRects(2) = WzRect_corners(WzRect_right(cutoutRect), WzRect_top(screenRect), WzRect_right(screenRect), WzRect_bottom(screenRect))
        ' bottom band below cutout rect
        surroundingRects(3) = WzRect_corners(WzRect_left(cutoutRect), WzRect_bottom(cutoutRect), WzRect_right(cutoutRect), WzRect_bottom(screenRect))

        For i = 0 To 3
            If surroundingRects(i).w <= 0 Or surroundingRects(i).h <= 0 Then
                GoTo continue_
            End If
            pie_UniTransBoxFill WzRect_left(surroundingRects(i)), WzRect_top(surroundingRects(i)), WzRect_right(surroundingRects(i)), WzRect_bottom(surroundingRects(i)), d.backgroundColor
continue_:
        Next
    End If
End Sub

Public Sub W_HELPSCREEN_CLICKFORM_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_HELPSCREEN_CLICKFORM
    Set d = W_HELPSCREEN_CLICKFORM_(self)
    If keyPressed(KEY_ESC) Then
        If Not d.onCancelPressed Is Nothing Then
            WzFunc_call d.onCancelPressed
        End If
    End If
    inputLoseFocus  ' clear the input buffer.
End Sub

Private Sub W_HELPSCREEN_CLICKFORM_setHelpForWidget(ByVal self As WIDGET, ByVal widget As WIDGET, ByVal widgetHelp As WidgetHelp)
    Dim d As W_HELPSCREEN_CLICKFORM, it As Variant, helpTriggerWidget As WIDGET
    If widget Is Nothing Then Exit Sub
    Set d = W_HELPSCREEN_CLICKFORM_(self)
    it = registeredWidgets_find(self, widget)
    If Not IsEmpty(it) Then
        ' update help
        d.registeredWidgets.Remove "p" & ObjPtr(widget)
        d.registeredWidgets.Add Array(widget, widgetHelp.copy(), it(2)), "p" & ObjPtr(widget)
        Exit Sub
    End If
    ' WidgetInfo info(widgetHelp);
    Set helpTriggerWidget = HelpTriggerWidget_make(self, widget)
    self.attach helpTriggerWidget, ChildZPos_Back
    d.registeredWidgets.Add Array(widget, widgetHelp.copy(), helpTriggerWidget), "p" & ObjPtr(widget)
End Sub

Private Sub W_HELPSCREEN_CLICKFORM_removeHelpForWidget(ByVal self As WIDGET, ByVal widget As WIDGET)
    Dim it As Variant, helpTriggerWidget As WIDGET
    If widget Is Nothing Then Exit Sub
    it = registeredWidgets_find(self, widget)
    If Not IsEmpty(it) Then
        Set helpTriggerWidget = it(2)
        helpTriggerWidget.deleteLater
        W_HELPSCREEN_CLICKFORM_(self).registeredWidgets.Remove "p" & ObjPtr(widget)
    End If
End Sub

' const WidgetHelp* getHelpForWidget(const std::shared_ptr<WIDGET>& widget) (Nothing = nullptr)
Private Function W_HELPSCREEN_CLICKFORM_getHelpForWidget(ByVal self As WIDGET, ByVal widget As WIDGET) As WidgetHelp
    Dim it As Variant
    it = registeredWidgets_find(self, widget)
    If IsEmpty(it) Then
        Set W_HELPSCREEN_CLICKFORM_getHelpForWidget = Nothing
        Exit Function
    End If
    Set W_HELPSCREEN_CLICKFORM_getHelpForWidget = it(1)
End Function

Private Sub W_HELPSCREEN_CLICKFORM_setCurrentHelpWidget(ByVal self As WIDGET, ByVal triggerWidget As WIDGET, ByVal helpPanel As WIDGET)
    Dim d As W_HELPSCREEN_CLICKFORM
    Set d = W_HELPSCREEN_CLICKFORM_(self)
    d.currentHelpWidget = HelpTriggerWidget_(triggerWidget).weakSerial
    If Not d.currentHelpPanel Is Nothing Then
        d.currentHelpPanel.deleteLater
    End If
    Set d.currentHelpPanel = helpPanel
    self.attach d.currentHelpPanel
    d.currentHelpPanel.callCalcLayout
End Sub

Private Sub W_HELPSCREEN_CLICKFORM_unsetCurrentHelpWidget(ByVal self As WIDGET, ByVal triggerWidget As WIDGET)
    Dim d As W_HELPSCREEN_CLICKFORM
    Set d = W_HELPSCREEN_CLICKFORM_(self)
    If WeakWidget_lock(d.currentHelpWidget) Is triggerWidget Then
        d.currentHelpWidget = 0                 ' currentHelpWidget.reset();
        If Not d.currentHelpPanel Is Nothing Then
            d.currentHelpPanel.deleteLater
            Set d.currentHelpPanel = Nothing
        End If
    End If
End Sub

Private Function W_HELPSCREEN_CLICKFORM_helperWidgetHasFocusLock(ByVal self As WIDGET) As Boolean
    Dim lockedScreen As W_SCREEN, widgetWithFocus As WIDGET
    Set lockedScreen = self.screenPointer
    If Not lockedScreen Is Nothing Then
        Set widgetWithFocus = lockedScreen.getWidgetWithFocus()
        W_HELPSCREEN_CLICKFORM_helperWidgetHasFocusLock = (Not widgetWithFocus Is Nothing And Not widgetWithFocus Is self)
        Exit Function
    End If
    W_HELPSCREEN_CLICKFORM_helperWidgetHasFocusLock = False
End Function

'--------------------------------------------------------------------------
' the lambdas

Public Function helpscreen_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psOverlayScreen As W_SCREEN, psWidget As WIDGET, width As Long, height As Long, triggerWidget As WIDGET
    Dim psParent As WIDGET, psTarget As WIDGET, targetedBoundingRect As WzRect, x0 As Long, y0 As Long, r As WzRect
    Select Case f.id
    Case FN_helpOverlay_rootClicked, FN_helpCloseButton_onClick ' [psWeakHelpOverlayScreen]() / (W_BUTTON& button)
        Set psOverlayScreen = helpScreen_lock(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            W_HELP_OVERLAY_SCREEN_closeHelpOverlayScreen psOverlayScreen
        End If

    Case FN_helpCloseButton_calcLayout          ' [minButtonWidthForText](WIDGET *psWidget)
        Set psWidget = args(0)
        width = f.cap(0) + CloseHelpPadding
        height = iV_GetTextLineSize(font_regular) + CloseHelpPadding
        psWidget.setGeometry (screenWidth - width) \ 2, CloseHelpMargin, width, height

    Case FN_helpTrigger_showTask                ' [triggerWidget]
        Set triggerWidget = f.cap(0)
        HelpTriggerWidget_showWidgetHelp triggerWidget

    Case FN_helpTrigger_closeTask               ' [triggerWidget]
        Set triggerWidget = f.cap(0)
        HelpTriggerWidget_closeWidgetHelp triggerWidget

    Case FN_helpDisplayForm_calcLayout          ' [weakParent](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = WeakWidget_lock(f.cap(0))
        If Not WIDGET_isKindOf(psParent, K_HelpTriggerWidget) Then Set psParent = Nothing
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/screens/helpscreen.cpp:782", "showWidgetHelp", "parent is null": Exit Function

        ' Get the bounding rect of the parent (trigger) widget, in screen coordinates
        targetedBoundingRect = psParent.screenGeometry()

        ' Try to get the position of the parent (target) widget as well
        Set psTarget = HelpTriggerWidget_(psParent).target
        If Not psTarget Is Nothing Then
            targetedBoundingRect = WzRect_minimumBoundingRect(targetedBoundingRect, psTarget.screenGeometry())
        End If

        ' Now try to position the help such that it is adjacent to this bounding rect, and aligned as best we can - but fully on-screen
        If psWidget.parent() Is Nothing Then
            ' skip setting - not attached to parent yet
            Exit Function
        End If

        ' Try to align with the left side of the targeted bounding rect
        x0 = targetedBoundingRect.x
        If x0 + psWidget.width() > screenWidth Then
            x0 = screenWidth - psWidget.width()
        End If

        ' Try to align so that it's "above" the targeted bounding rect
        y0 = targetedBoundingRect.y - (psWidget.height() + 1)
        If y0 < 0 Then
            ' Would be off-screen
            ' Try to align so that it's "below" the targeted bounding rect
            y0 = WzRect_bottom(targetedBoundingRect) + 1
            If y0 + psWidget.height() > screenHeight Then
                ' Would be off-screen
                ' Just make sure it's on-screen as best we can
                y0 = screenHeight - psWidget.height()
                If y0 < 0 Then y0 = 0
            End If
        End If

        r = WzRect_(x0, y0, psWidget.width(), psWidget.height())
        psWidget.setGeometryFromScreenRect r

    Case FN_helpClickForm_calcLayout            ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth, screenHeight
    End Select
End Function
